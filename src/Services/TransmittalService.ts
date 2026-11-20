/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

/**
 * Send-Transmittal upload service.
 *
 * Every ticked document is copied into "ProjectTransmittals" and stamped with a COPY
 * of the source document's own metadata. The source document is never moved or modified.
 *
 * Source column   ->  target column on the uploaded copy
 * ------------------------------------------------------------------------------------
 * ReferenceNo     ->  ReferenceNo      (falls back to the table row value)
 * Active          ->  Active           ("1" unless the source column explicitly says No)
 * Publisher       ->  Publisher        (person, resolved from the source PublisherId)
 * PublisherEmail  ->  PublisherEmail   (source value, else the publishers' e-mails)
 * Status          ->  Status           (source lookup id, else "Published")
 * InternalStatus  ->  InternalStatus   (falls back to the table row value)
 * Level           ->  Level            (falls back to the table row value)
 * ActualName      ->  ActualName       (falls back to the table row, then the file name)
 * (popup Remarks) ->  Description
 *
 * NOTE (column assumptions - adjust here only if a column is renamed):
 *  Description, ReferenceNo, Active, Publisher, PublisherEmail, Status,
 *  InternalStatus, Level, ActualName.
 * "Status" is a lookup column whose source list (DMS_Mas_Status) provides the fallback
 * status TRANSMITTAL_FALLBACK_STATUS_NAME ("Published").
 */

/** Library that receives the transmittal copies. */
export const TRANSMITTAL_LIBRARY_NAME = "ProjectTransmittals";

/**
 * Status used when the source item has no Status value, or for every copy when
 * TRANSMITTAL_KEEP_SOURCE_STATUS is false.
 */
export const TRANSMITTAL_FALLBACK_STATUS_NAME = "Published";

/**
 * true  -> each copy keeps the source item's Status lookup value;
 * false -> every copy receives TRANSMITTAL_FALLBACK_STATUS_NAME ("Published").
 */
export const TRANSMITTAL_KEEP_SOURCE_STATUS = true;

/** Yes/No metadata values understood by ValidateUpdateListItem ("1" = Yes, "0" = No). */
const TRANSMITTAL_ACTIVE_VALUE = "1";
const TRANSMITTAL_INACTIVE_VALUE = "0";

export interface ITransmittalDocument {
    /** Stable row key (list item id). */
    id: string;
    /** Physical file name including its extension. */
    name: string;
    /** Display name (ActualName when available). */
    actualName: string;
    /** Server-relative URL of the source file. */
    serverRelativeUrl: string;
    referenceNo?: string;
    /** Raw level / version value as shown in the document table. */
    version?: string;
    /** Friendly status (DisplayStatus first). */
    status?: string;
    internalStatus?: string;
    level?: string;
}

export interface ITransmittalDocumentResult {
    document: ITransmittalDocument;
    success: boolean;
    uploadedFileName?: string;
    uploadedServerRelativeUrl?: string;
    error?: string;
}

export interface IUploadTransmittalOptions {
    documents: ITransmittalDocument[];
    remarks?: string;
    /**
     * Optional transmittal number. It is only used to build the unique file name
     * prefix; there is no transmittal-number source in the current codebase, so a
     * local timestamp is used when it is not supplied.
     */
    transmittalNumber?: string;
}

export interface ISendTransmittalForm {
    clientEmails: string[];
    subject: string;
    remarks: string;
}

interface IMasStatusItem {
    Id: number;
    Title?: string;
    StatusName?: string;
    InternalStatus?: string;
}

/** A SharePoint user resolved from a person column value. */
interface ITransmittalUser {
    loginName: string;
    email: string;
}

/** Per-run caches - no lookup is requested twice while one transmittal is sent. */
interface ITransmittalRunCache {
    /** Fallback Status lookup id, resolved lazily and only once per run. */
    fallbackStatusId: Promise<string> | null;
    /** SharePoint user id -> resolved user (null = could not be resolved). */
    usersById: { [userId: string]: ITransmittalUser | null };
}

const normalizeStatusValue = (value?: string): string => String(value || "").trim().toLowerCase();

/** Error thrown when at least one document could not be added. */
export class TransmittalUploadError extends Error {
    public readonly results: ITransmittalDocumentResult[];
    public readonly failedCount: number;
    public readonly totalCount: number;

    public constructor(failedCount: number, totalCount: number, firstReason: string, results: ITransmittalDocumentResult[]) {
        super(`${failedCount} of ${totalCount} document(s) could not be added to ${TRANSMITTAL_LIBRARY_NAME}. First error: ${firstReason}`);
        this.name = "TransmittalUploadError";
        this.failedCount = failedCount;
        this.totalCount = totalCount;
        this.results = results;
        Object.setPrototypeOf(this, TransmittalUploadError.prototype);
    }
}

/**
 * Maps a ReusableDataTable row (SPO list item) to the data the transmittal needs.
 * Missing values are tolerated - the corresponding metadata is simply left blank.
 */
export const toTransmittalDocument = (row: any): ITransmittalDocument => {
    const fields: any = row?.ListItemAllFields || {};
    const physicalName = String(row?.Name || fields?.FileLeafRef || fields?.ActualName || "");
    const actualName = String(fields?.ActualName || physicalName);
    const extension = physicalName.indexOf(".") > -1 ? physicalName.split(".").pop()!.toLowerCase() : "";

    // Same version rule the document table uses (PDF shows Level, others the file version).
    const rawVersion = extension === "pdf" ? fields?.Level : fields?.OData__UIVersionString;

    return {
        id: String(row?.id ?? row?.ID ?? fields?.ID ?? physicalName),
        name: physicalName,
        actualName,
        serverRelativeUrl: String(row?.ServerRelativeUrl || fields?.File?.ServerRelativeUrl || ""),
        referenceNo: fields?.ReferenceNo ? String(fields.ReferenceNo) : "",
        version: rawVersion ? String(rawVersion) : "",
        status: fields?.DisplayStatus ? String(fields.DisplayStatus) : (fields?.Status?.StatusName ? String(fields.Status.StatusName) : ""),
        internalStatus: fields?.InternalStatus ? String(fields.InternalStatus) : "",
        level: fields?.Level ? String(fields.Level) : ""
    };
};

/** Prefixes the table version with "v" unless it already starts with it. */
export const formatTransmittalVersion = (version?: string): string => {
    const value = (version || "").trim();
    if (!value) {
        return "";
    }
    return /^v/i.test(value) ? value : `v${value}`;
};

/** Removes characters SharePoint rejects in file names. */
export const sanitizeSharePointFileName = (fileName: string): string => {
    return String(fileName || "")
        .replace(/[~"#%&*:<>?\/\\{|}]/g, "_")
        .replace(/[\u0000-\u001F]/g, "")
        .trim()
        .replace(/[. ]+$/, "");
};

/**
 * Builds a unique, sanitized file name for the copy.
 * Example: "20261116-142530_Basement Layout.docx"
 */
export const buildTransmittalFileName = (originalName: string, transmittalNumber?: string, now: Date = new Date()): string => {
    const sanitized = sanitizeSharePointFileName(originalName) || "document";
    const prefix = transmittalNumber && transmittalNumber.trim()
        ? sanitizeSharePointFileName(transmittalNumber)
        : formatTimestamp(now);

    return `${prefix}_${sanitized}`;
};

/**
 * Copies every document into "ProjectTransmittals", writes the transmittal
 * metadata and returns a per-document result.
 *
 * One request digest is requested per run. A failure on one document does not
 * stop the remaining documents, but the returned promise rejects with a
 * TransmittalUploadError when anything failed (its `results` still contains the
 * per-document outcome). Callers should only continue to the cover sheet / email
 * step when the promise resolves.
 */
export async function uploadTransmittalDocuments(
    context: WebPartContext,
    options: IUploadTransmittalOptions
): Promise<ITransmittalDocumentResult[]> {
    const documents = options?.documents || [];
    if (documents.length === 0) {
        return [];
    }

    const webUrl = context.pageContext.web.absoluteUrl.replace(/\/+$/, "");
    const loginName = context.pageContext.user.loginName || "";
    const userEmail = context.pageContext.user.email || "";

    if (!loginName) {
        throw new Error("The current user's login name is not available, so the Publisher column cannot be set.");
    }

    // One digest for the whole run (uploads + metadata + rollback).
    const digest = await getRequestDigest(webUrl);
    const rootFolderUrl = await getLibraryRootFolder(webUrl);
    const runDate = new Date();

    const currentUser: ITransmittalUser = { loginName, email: userEmail };
    /** Status / user lookups are resolved lazily and at most once per run. */
    const runCache: ITransmittalRunCache = { fallbackStatusId: null, usersById: {} };

    const results: ITransmittalDocumentResult[] = [];
    /* Copies created by this run - removed again if any document fails. */
    const uploadedServerRelativeUrls: string[] = [];

    for (const document of documents) {
        let uploadedServerRelativeUrl = "";

        try {
            const sourceUrl = toAbsoluteUrl(webUrl, document.serverRelativeUrl);
            if (!sourceUrl) {
                throw new Error("The source file URL is missing.");
            }

            // Metadata first: reading the source item and resolving the publisher can fail,
            // and a failure there must never leave an uploaded copy behind.
            const sourceFields = await readSourceItemFields(webUrl, document);
            const metadata = await buildTransmittalMetadata(
                webUrl,
                document,
                sourceFields,
                currentUser,
                runCache,
                options.remarks || ""
            );

            const sourceBlob = await readSourceFile(sourceUrl);
            const uploadedFileName = buildTransmittalFileName(document.name, options.transmittalNumber, runDate);

            uploadedServerRelativeUrl = await uploadCopy(webUrl, rootFolderUrl, uploadedFileName, sourceBlob, digest);
            const itemId = await getFileItemId(webUrl, uploadedServerRelativeUrl);

            await applyTransmittalMetadata(webUrl, itemId, metadata, digest);

            results.push({
                document,
                success: true,
                uploadedFileName,
                uploadedServerRelativeUrl
            });
            // Only fully tagged copies are kept; they are rolled back below if
            // a later document fails so a retry cannot create duplicates.
            uploadedServerRelativeUrls.push(uploadedServerRelativeUrl);
        } catch (error) {
            // Never leave an un-tagged copy behind.
            if (uploadedServerRelativeUrl) {
                await deleteFileQuietly(webUrl, uploadedServerRelativeUrl, digest);
            }

            results.push({
                document,
                success: false,
                uploadedFileName: undefined,
                uploadedServerRelativeUrl: undefined,
                error: `${document.actualName || document.name}: ${getErrorMessage(error)}`
            });
        }
    }

    const failures = results.filter((result) => !result.success);
    if (failures.length > 0) {
        // All-or-nothing: drop every copy this run created (best effort) so the
        // retry starts clean. The failing document already removed its own copy.
        for (const uploadedUrl of uploadedServerRelativeUrls) {
            await deleteFileQuietly(webUrl, uploadedUrl, digest);
        }

        throw new TransmittalUploadError(failures.length, documents.length, failures[0].error || "Unknown error", results);
    }

    return results;
}

/* -------------------------------------------------------------------------- */
/* SharePoint REST helpers                                                     */
/* -------------------------------------------------------------------------- */

const pad = (value: number): string => (value < 10 ? `0${value}` : String(value));

const formatTimestamp = (date: Date): string =>
    `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}-` +
    `${pad(date.getHours())}${pad(date.getMinutes())}${pad(date.getSeconds())}`;

const escapeODataString = (value: string): string => String(value || "").replace(/'/g, "''");

/**
 * Encodes a value that is placed inside a quoted REST segment, i.e.
 * `decodedurl='...'` / `url='...'`. A single quote is doubled BEFORE encoding
 * so a name like "John's report.docx" no longer produces an invalid URL.
 */
const toQuotedUrlValue = (value: string): string =>
    encodeURIComponent(String(value || "").replace(/'/g, "''"));

const getErrorMessage = (error: any): string => {
    if (!error) {
        return "Unknown error";
    }
    if (typeof error === "string") {
        return error;
    }
    return error.message || String(error);
};

const toAbsoluteUrl = (webUrl: string, serverRelativeUrl: string): string => {
    if (!serverRelativeUrl) {
        return "";
    }
    if (/^https?:\/\//i.test(serverRelativeUrl)) {
        return serverRelativeUrl;
    }
    try {
        return new URL(serverRelativeUrl, webUrl).toString();
    } catch {
        return "";
    }
};

/**
 * Minimal response surface used by the helpers below.
 *
 * Both the browser `Response` (returned by `fetch`) and `SPHttpClientResponse`
 * (returned by `spHttpClient`) satisfy it, so one helper can serve both call paths.
 * `Response` itself cannot be used here: SPFx 1.22 ships `SPHttpClientResponse`
 * (via `HttpClientResponse implements Response`) whose members no longer match the
 * current `lib.dom` `Response` definition (e.g. `bytes()`, `body`, `redirected`).
 */
interface ISharePointResponse {
    readonly ok: boolean;
    readonly status: number;
    json(): Promise<any>;
    text(): Promise<string>;
}

const readErrorMessage = async (response: ISharePointResponse, fallback: string): Promise<string> => {
    let text = "";
    try {
        text = await response.text();
    } catch {
        return `${fallback} (HTTP ${response.status})`;
    }

    if (text) {
        try {
            const data = JSON.parse(text);
            const message =
                data?.["odata.error"]?.message?.value ||
                data?.error?.message?.value ||
                (typeof data?.error?.message === "string" ? data.error.message : "") ||
                (typeof data?.error_description === "string" ? data.error_description : "");
            if (message) {
                return message;
            }
        } catch {
            // Not JSON - fall through to the raw text.
        }

        if (text.length <= 300 && text.indexOf("<html") === -1) {
            return text;
        }
    }

    return `${fallback} (HTTP ${response.status})`;
};

/** Reads a JSON body from a `fetch` or `spHttpClient` response; null when it is empty/invalid. */
const readJson = async (response: ISharePointResponse): Promise<any> => {
    try {
        return await response.json();
    } catch {
        return null;
    }
};

const getRequestDigest = async (webUrl: string): Promise<string> => {
    const response = await fetch(`${webUrl}/_api/contextinfo`, {
        method: "POST",
        credentials: "same-origin",
        headers: {
            Accept: "application/json;odata=verbose"
        }
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, "Unable to request a SharePoint request digest."));
    }

    const data = await readJson(response);
    const digest = data?.d?.GetContextWebInformation?.FormDigestValue || data?.FormDigestValue;
    if (!digest) {
        throw new Error("SharePoint did not return a request digest.");
    }

    return String(digest);
};

const getLibraryRootFolder = async (webUrl: string): Promise<string> => {
    const url = `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(TRANSMITTAL_LIBRARY_NAME)}')` +
        `?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;

    const response = await fetch(url, {
        credentials: "same-origin",
        headers: {
            Accept: "application/json;odata=nometadata"
        }
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, `The library "${TRANSMITTAL_LIBRARY_NAME}" could not be found.`));
    }

    const data = await readJson(response);
    const rootFolder = data?.RootFolder?.ServerRelativeUrl;
    if (!rootFolder) {
        throw new Error(`The library "${TRANSMITTAL_LIBRARY_NAME}" does not expose a root folder.`);
    }

    return String(rootFolder).replace(/\/+$/, "");
};



/**
 * Reads every item of the Status lookup source list (DMS_Mas_Status), following
 * paging. Some tenants name the columns differently, so the select list degrades
 * gracefully (400 = unknown column) instead of failing the whole transmittal.
 */
const readStatusSourceItems = async (webUrl: string, lookupListId: string): Promise<IMasStatusItem[]> => {
    const itemsUrl = `${webUrl}/_api/web/lists(guid'${escapeODataString(lookupListId)}')/items`;
    const selectVariants = ["Id,Title,StatusName,InternalStatus", "Id,Title", ""];

    for (const select of selectVariants) {
        const statuses: IMasStatusItem[] = [];
        let nextUrl: string | undefined = `${itemsUrl}?${select ? `$select=${select}&` : ""}$top=500`;
        let unusable = false;

        for (let page = 0; nextUrl && page < 20; page++) {
            const response: Response = await fetch(nextUrl, {
                credentials: "same-origin",
                headers: { Accept: "application/json;odata=nometadata" }
            });

            // 400 means one of the selected columns does not exist -> retry with fewer.
            if (response.status === 400) {
                unusable = true;
                break;
            }

            if (!response.ok) {
                throw new Error(await readErrorMessage(response, "The statuses could not be read from DMS_Mas_Status."));
            }

            const data = await readJson(response);
            statuses.push(...((data?.value as IMasStatusItem[]) || []));
            nextUrl = data?.["odata.nextLink"] || undefined;
        }

        if (!unusable) {
            return statuses;
        }
    }

    throw new Error("The statuses could not be read from DMS_Mas_Status.");
};

/**
 * Resolves the id of TRANSMITTAL_FALLBACK_STATUS_NAME ("Published") in the Status
 * lookup source list. No id is hard-coded: every item is read and matched on
 * StatusName first, then InternalStatus, then Title (case and extra spaces ignored).
 */
const resolveFallbackStatusId = async (webUrl: string): Promise<string> => {
    const fieldUrl = `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(TRANSMITTAL_LIBRARY_NAME)}')` +
        `/fields/getbyinternalnameortitle('Status')?$select=InternalName,Title,TypeAsString,LookupList`;

    const fieldResponse = await fetch(fieldUrl, {
        credentials: "same-origin",
        headers: { Accept: "application/json;odata=nometadata" }
    });

    if (!fieldResponse.ok) {
        throw new Error(await readErrorMessage(fieldResponse, `The "Status" column could not be read from "${TRANSMITTAL_LIBRARY_NAME}".`));
    }

    const field = await readJson(fieldResponse);
    if (String(field?.TypeAsString || "") !== "Lookup") {
        throw new Error(`The "Status" column in "${TRANSMITTAL_LIBRARY_NAME}" is not a lookup column (found "${field?.TypeAsString}").`);
    }

    // LookupList may be returned with braces, e.g. "{6f2b...}".
    const lookupListId = String(field?.LookupList || "").replace(/[{}]/g, "");
    if (!lookupListId) {
        throw new Error(`The "Status" column in "${TRANSMITTAL_LIBRARY_NAME}" has no lookup source list.`);
    }

    // TODO(transmittal): if the lookup source lives on another web, replace this with openWebById().
    const statuses = await readStatusSourceItems(webUrl, lookupListId);
    const target = normalizeStatusValue(TRANSMITTAL_FALLBACK_STATUS_NAME);
    const columns: Array<"StatusName" | "InternalStatus" | "Title"> = ["StatusName", "InternalStatus", "Title"];

    for (const column of columns) {
        const match = statuses.find((status) => normalizeStatusValue(status[column]) === target);
        if (match?.Id) {
            return String(match.Id);
        }
    }

    const available = statuses
        .map((status) => status.StatusName || status.InternalStatus || status.Title)
        .filter(Boolean)
        .join(", ");
    throw new Error(
        `The status "${TRANSMITTAL_FALLBACK_STATUS_NAME}" was not found in DMS_Mas_Status. Available values: ${available || "none"}.`
    );
};

/**
 * Fallback Status id, resolved lazily and only once per run: documents that still
 * carry their own Status never trigger the extra request.
 */
const getFallbackStatusId = async (webUrl: string, cache: ITransmittalRunCache): Promise<string> => {
    if (!cache.fallbackStatusId) {
        cache.fallbackStatusId = resolveFallbackStatusId(webUrl);
    }

    return await cache.fallbackStatusId;
};

/** Turns a server-relative or absolute file URL into a decoded server-relative path. */
const toServerRelativePath = (serverRelativeUrl: string): string => {
    const value = String(serverRelativeUrl || "").trim();
    if (!value) {
        return "";
    }

    if (!/^https?:\/\//i.test(value)) {
        return value.charAt(0) === "/" ? value : `/${value}`;
    }

    try {
        const path = new URL(value).pathname;
        try {
            return decodeURIComponent(path);
        } catch {
            return path;
        }
    } catch {
        return "";
    }
};

/**
 * Reads the SOURCE document's own list item (ListItemAllFields).
 * Returns null when the item cannot be read: the caller then keeps the values that
 * are already on the table row, so a locked or checked-out file still works.
 */
const readSourceItemFields = async (webUrl: string, document: ITransmittalDocument): Promise<any | null> => {
    const serverRelativePath = toServerRelativePath(document.serverRelativeUrl);
    if (!serverRelativePath) {
        return null;
    }

    const url = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${toQuotedUrlValue(serverRelativePath)}')` +
        `/ListItemAllFields`;

    try {
        const response = await fetch(url, {
            credentials: "same-origin",
            headers: { Accept: "application/json;odata=nometadata" }
        });

        if (!response.ok) {
            return null;
        }

        const data = await readJson(response);
        return data?.d || data || null;
    } catch {
        // Reading the source metadata is best effort - fall back to the table row.
        return null;
    }
};

/** Person/lookup id values arrive as a number, a string, or an array (multi-value). */
const toIdList = (value: any): string[] => {
    if (value === undefined || value === null) {
        return [];
    }

    const values: any[] = Array.isArray(value)
        ? value
        : (value && Array.isArray(value.results) ? value.results : [value]);

    return values
        .map((entry) => String(entry === undefined || entry === null ? "" : entry).trim())
        .filter((entry) => entry.length > 0);
};

/**
 * Yes/No value for the copy: "0" only when the source column explicitly says No,
 * otherwise "1" (also when the source value is empty or the column is missing).
 */
const toActiveValue = (value: any): string => {
    if (value === false || value === 0) {
        return TRANSMITTAL_INACTIVE_VALUE;
    }

    const text = String(value === undefined || value === null ? "" : value).trim().toLowerCase();
    if (!text) {
        return TRANSMITTAL_ACTIVE_VALUE;
    }

    return text === "false" || text === "0" || text === "no"
        ? TRANSMITTAL_INACTIVE_VALUE
        : TRANSMITTAL_ACTIVE_VALUE;
};

const readSourceFile = async (sourceUrl: string): Promise<Blob> => {
    const response = await fetch(sourceUrl, {
        credentials: "same-origin"
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, "The source file could not be read."));
    }

    return await response.blob();
};

/**
 * Resolves a person column id to its login name and e-mail.
 * Results are cached for the run; users that cannot be resolved are skipped.
 */
const resolveUserById = async (
    webUrl: string,
    userId: string,
    cache: ITransmittalRunCache
): Promise<ITransmittalUser | null> => {
    if (!/^\d+$/.test(userId)) {
        return null;
    }

    if (Object.prototype.hasOwnProperty.call(cache.usersById, userId)) {
        return cache.usersById[userId];
    }

    let user: ITransmittalUser | null = null;
    try {
        const response = await fetch(`${webUrl}/_api/web/getuserbyid(${userId})?$select=LoginName,Email`, {
            credentials: "same-origin",
            headers: { Accept: "application/json;odata=nometadata" }
        });

        if (response.ok) {
            const data = await readJson(response);
            const payload = data?.d || data || {};
            const loginName = String(payload.LoginName || "").trim();
            const email = String(payload.Email || "").trim();
            if (loginName) {
                user = { loginName, email };
            }
        }
    } catch {
        user = null;
    }

    cache.usersById[userId] = user;
    return user;
};

/**
 * Builds the metadata of one copy from the SOURCE item, falling back to the table
 * row. Called before the upload so a metadata problem never leaves a copy behind.
 */
const buildTransmittalMetadata = async (
    webUrl: string,
    document: ITransmittalDocument,
    sourceFields: any,
    currentUser: ITransmittalUser,
    cache: ITransmittalRunCache,
    remarks: string
): Promise<ITransmittalMetadata> => {
    const source = sourceFields || {};
    const textValue = (value: any): string => (value === undefined || value === null ? "" : String(value).trim());

    // Status: keep the source lookup value unless it is empty (or the switch is off).
    const sourceStatusId = TRANSMITTAL_KEEP_SOURCE_STATUS ? textValue(source.StatusId) : "";
    const statusLookupId = sourceStatusId || await getFallbackStatusId(webUrl, cache);

    // Publisher: the source item's PublisherId, resolved to login names; the current
    // user is used when the source has no (resolvable) publisher.
    const publishers: ITransmittalUser[] = [];
    for (const userId of toIdList(source.PublisherId)) {
        const user = await resolveUserById(webUrl, userId, cache);
        if (user) {
            publishers.push(user);
        }
    }
    if (publishers.length === 0) {
        publishers.push(currentUser);
    }

    const publisherEmails = publishers
        .map((user) => user.email)
        .filter((email) => email.length > 0)
        .join("; ");

    return {
        remarks: remarks || "",
        referenceNo: textValue(source.ReferenceNo) || document.referenceNo || "",
        active: toActiveValue(source.Active),
        publisherLoginNames: publishers
            .map((user) => user.loginName)
            .filter((loginName) => loginName.length > 0),
        publisherEmail: textValue(source.PublisherEmail) || publisherEmails || currentUser.email,
        statusLookupId,
        internalStatus: textValue(source.InternalStatus) || document.internalStatus || "",
        level: textValue(source.Level) || document.level || "",
        actualName: textValue(source.ActualName) || document.actualName || document.name
    };
};

const uploadCopy = async (
    webUrl: string,
    rootFolderUrl: string,
    fileName: string,
    fileBlob: Blob,
    digest: string
): Promise<string> => {
    const url = `${webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${toQuotedUrlValue(rootFolderUrl)}')` +
        `/Files/add(overwrite=false,url='${toQuotedUrlValue(fileName)}')`;

    const response = await fetch(url, {
        method: "POST",
        credentials: "same-origin",
        headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/octet-stream",
            "X-RequestDigest": digest
        },
        body: fileBlob
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, `"${fileName}" could not be copied to ${TRANSMITTAL_LIBRARY_NAME}.`));
    }

    const data = await readJson(response);
    const uploadedUrl = data?.ServerRelativeUrl || data?.d?.ServerRelativeUrl || `${rootFolderUrl}/${fileName}`;
    return String(uploadedUrl);
};

const getFileItemId = async (webUrl: string, fileServerRelativeUrl: string): Promise<number> => {
    const url = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${toQuotedUrlValue(fileServerRelativeUrl)}')` +
        `/ListItemAllFields?$select=Id`;

    const response = await fetch(url, {
        credentials: "same-origin",
        headers: {
            Accept: "application/json;odata=nometadata"
        }
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, "The uploaded copy could not be located."));
    }

    const data = await readJson(response);
    const itemId = data?.Id ?? data?.d?.Id;
    if (itemId === undefined || itemId === null) {
        throw new Error("The uploaded copy has no list item id.");
    }

    return Number(itemId);
};

interface ITransmittalMetadata {
    remarks: string;
    referenceNo: string;
    /** "1" / "0" - copied from the source item's Active column. */
    active: string;
    /** Login names of the source publishers (the current user when there are none). */
    publisherLoginNames: string[];
    publisherEmail: string;
    statusLookupId: string;
    internalStatus: string;
    level: string;
    actualName: string;
}

const applyTransmittalMetadata = async (
    webUrl: string,
    itemId: number,
    metadata: ITransmittalMetadata,
    digest: string
): Promise<void> => {
    const formValues = [
        { FieldName: "Remark", FieldValue: metadata.remarks },
        { FieldName: "ReferenceNo", FieldValue: metadata.referenceNo },
        { FieldName: "Active", FieldValue: metadata.active },
        // TODO(transmittal): a single-value Publisher column only accepts the first entry.
        { FieldName: "Publisher", FieldValue: JSON.stringify(metadata.publisherLoginNames.map((loginName) => ({ Key: loginName }))) },
        { FieldName: "PublisherEmail", FieldValue: metadata.publisherEmail },
        { FieldName: "Status", FieldValue: metadata.statusLookupId },
        { FieldName: "InternalStatus", FieldValue: metadata.internalStatus },
        { FieldName: "Level", FieldValue: metadata.level },
        { FieldName: "ActualName", FieldValue: metadata.actualName }
    ];

    const url = `${webUrl}/_api/web/lists/getbytitle('${escapeODataString(TRANSMITTAL_LIBRARY_NAME)}')` +
        `/items(${itemId})/validateupdatelistitem`;

    const response = await fetch(url, {
        method: "POST",
        credentials: "same-origin",
        headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest
        },
        body: JSON.stringify({ formValues, bNewDocumentUpdate: false })
    });

    if (!response.ok) {
        throw new Error(await readErrorMessage(response, "The transmittal metadata could not be saved."));
    }

    const data = await readJson(response);
    const fieldResults: any[] = data?.value || data?.d?.ValidateUpdateListItem?.results || [];
    const failedField = fieldResults.filter((result) => result && result.HasException)[0];

    if (failedField) {
        const fieldName = failedField.FieldName ? ` (${failedField.FieldName})` : "";
        throw new Error(failedField.ErrorMessage || `A metadata value was rejected${fieldName}.`);
    }
};

/* ------------------------------------------------------------------ */
/* Read: transmittal audit log                                         */
/* ------------------------------------------------------------------ */

/**
 * Reads every copy in "ProjectTransmittals" (one item per document).
 *
 * The send flow writes all copies to the library root (see `uploadCopy`), so an
 * item carries no "source folder" column. The caller therefore scopes the rows
 * to the open folder by matching the copy name against the files on screen
 * (see `mapAuditRow` / `filterAuditLogEntries` in TransmittalAuditLog.tsx) and
 * sorts them by `Created`, newest first.
 */
export async function getTransmittalAuditLog(context: WebPartContext): Promise<any[]> {
    const url = `${context.pageContext.web.absoluteUrl}` +
        `/_api/web/lists/getbytitle('${escapeODataString(TRANSMITTAL_LIBRARY_NAME)}')/items` +
        `?$select=*,Author/Title,File/Name&$expand=Author,File&$orderby=Id desc&$top=5000`;

    const response = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: {
            Accept: "application/json;odata=nometadata",
            "odata-version": ""
        }
    });

    if (!response.ok) {
        throw new Error(`The transmittal audit log could not be loaded (${response.status}).`);
    }

    const data = await readJson(response);
    const items = data?.value || data?.d?.results || [];
    return items;
}

const deleteFileQuietly = async (webUrl: string, fileServerRelativeUrl: string, digest: string): Promise<void> => {
    try {
        const url = `${webUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${toQuotedUrlValue(fileServerRelativeUrl)}')`;
        await fetch(url, {
            method: "POST",
            credentials: "same-origin",
            headers: {
                Accept: "application/json;odata=nometadata",
                "X-RequestDigest": digest,
                "X-HTTP-Method": "DELETE",
                "IF-MATCH": "*"
            }
        });
    } catch {
        // Best effort only - the original error is the one worth reporting.
    }
};
