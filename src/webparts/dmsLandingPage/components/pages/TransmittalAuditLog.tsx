/* eslint-disable */
import * as React from "react";
import { useCallback, useEffect, useMemo, useState } from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getTransmittalAuditLog } from "../../../../Services/TransmittalService";

/**
 * Transmittal Audit Log
 * ---------------------------------------------------------------------------
 * Two independent pieces that share one design language:
 *
 *  1. <TransmittalAuditLogToggle /> - the "Audit Log" pill that lives in the
 *     workspace action bar (blue badge = number of entries that belong to the
 *     documents currently listed).
 *  2. <TransmittalAuditLog />       - the collapsible panel that slides open
 *     below the action bar.
 *
 * The panel is a controlled component: the entries array and the open/closed
 * flag stay in the page (React state), exactly like the rest of the workspace.
 *
 * Notes on the requested stack
 * ---------------------------------------------------------------------------
 * This SPFx solution does not ship Tailwind CSS or lucide-react (see
 * package.json), and adding a new CSS pipeline to the Heft build would risk the
 * existing bundle. The requested design tokens are therefore implemented as
 * scoped, hand-written CSS (`tal-` prefix, injected once) and the icons are
 * inline SVGs using lucide's own path data. Both look identical to the
 * Tailwind/lucide version at runtime and add no new dependency.
 */

/* ------------------------------------------------------------------ */
/* Lucide icons (same path data as lucide-react, no package needed)    */
/* ------------------------------------------------------------------ */
type ILucideIconProps = React.SVGProps<SVGSVGElement>;

const LucideIcon: React.FC<ILucideIconProps & { children: React.ReactNode }> = ({ children, ...props }) => (
    <svg
        xmlns="http://www.w3.org/2000/svg"
        viewBox="0 0 24 24"
        width="24"
        height="24"
        fill="none"
        stroke="currentColor"
        strokeWidth={2}
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
        focusable="false"
        {...props}
    >
        {children}
    </svg>
);

/** lucide-react `FileText`. */
const FileTextIcon: React.FC<ILucideIconProps> = (props) => (
    <LucideIcon {...props}>
        <path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z" />
        <path d="M14 2v4a2 2 0 0 0 2 2h4" />
        <path d="M10 9H8" />
        <path d="M16 13H8" />
        <path d="M16 17H8" />
    </LucideIcon>
);

/* ------------------------------------------------------------------ */
/* Public types                                                        */
/* ------------------------------------------------------------------ */
export interface ITransmittalAuditLogDocument {
    /** Optional list item id - lets an entry be matched by id as well as name. */
    id?: string;
    /** File name as it was sent (e.g. "Invoice_March.docx"). */
    name: string;
    /**
     * Extra spellings of the same file (Name / ActualName / FileLeafRef) used
     * only for folder matching. Never rendered.
     */
    matches?: string[];
}

/** One row of the transmittal audit log. Newest entries come first. */
export interface ITransmittalAuditLogEntry {
    /** Unique row key. */
    id: string;
    /** List item id of the primary document (optional, matched by name otherwise). */
    documentId?: string;
    /** Human readable transmittal number, e.g. "TR-2026-0001". */
    transmittalNo: string;
    /** Display name of the sender. */
    sentBy: string;
    /** Small subtitle under the sender ("Project Manager"). */
    role: string;
    /** Already formatted stamp ("18 Nov 2026, 11:56 AM"). */
    dateTime: string;
    /** Recipient e-mail address(es). */
    recipient: string;
    /** Free-text remark of the transmittal ("" when none). */
    remark?: string;
    documents: ITransmittalAuditLogDocument[];
    /** Full folder path, e.g. "Projects / Gulf Helicopters / Accounts Payable". */
    folderPath: string;
    /** Sort key: the service `Created` value as epoch ms (0 when unknown). */
    createdTime?: number;
}

export interface ITransmittalAuditLogToggleProps {
    /** Whether the panel is currently expanded. */
    isOpen: boolean;
    /** Number of transmittals - rendered in the blue badge. */
    count: number;
    onToggle: () => void;
}

export interface ITransmittalAuditLogProps {
    isOpen: boolean;
    /** Web part context - the card reads "ProjectTransmittals" with it. */
    context: WebPartContext;
    /**
     * Files visible in the open folder. The log is scoped to them because the
     * copies in "ProjectTransmittals" do not store the folder they came from.
     */
    documents?: IAuditLogDocumentRef[];
    /**
     * Path of the open folder. Only used as a fetch key: switching folders
     * reloads the rows, while opening/closing the card never refetches.
     */
    folderPath?: string;
    /** Bumped by the page after a transmittal was sent - forces a refetch. */
    refreshToken?: number;
    /** Reports the number of transmittals currently visible (blue badge). */
    onCountChange?: (count: number) => void;
    /** Optional "Cover" action. The button is always shown for layout parity. */
    onViewCover?: (entry: ITransmittalAuditLogEntry) => void;
    /** Optional Escape-to-close handler wired by the page that owns `isOpen`. */
    onRequestClose?: () => void;
}

/* ------------------------------------------------------------------ */
/* Current-view filtering                                              */
/* ------------------------------------------------------------------ */
/** Minimal shape of a document that is visible in the open folder. */
export interface IAuditLogDocumentRef {
    /** List item id (string form) when available. */
    id?: string;
    /** File name as shown in the grid ("Invoice_March.docx"). */
    name: string;
}

/**
 * Keeps only the audit entries that belong to at least one document of the
 * current view. Matching is done on the list item id first and on the file
 * name second, so entries stay visible even when the id is not known.
 */
export const filterAuditLogEntries = (
    entries: ITransmittalAuditLogEntry[],
    documentsInView: IAuditLogDocumentRef[]
): ITransmittalAuditLogEntry[] => {
    if (!documentsInView || documentsInView.length === 0) {
        return [];
    }

    /*
     * Plain lookup objects instead of Set/Map: the SPFx bundle targets ES5 and
     * this keeps the helper free of collection shims.
     */
    const ids: { [key: string]: boolean } = {};
    const names: { [key: string]: boolean } = {};

    documentsInView.forEach((document) => {
        if (document.id) {
            ids[String(document.id).toLowerCase()] = true;
        }
        if (document.name) {
            names[String(document.name).toLowerCase()] = true;
        }
    });

    const isCurrentViewDocument = (id?: string, name?: string): boolean =>
        (!!id && ids[String(id).toLowerCase()] === true) ||
        (!!name && names[String(name).toLowerCase()] === true);

    const matchesCurrentView = (document: ITransmittalAuditLogDocument): boolean =>
        isCurrentViewDocument(document.id, document.name) ||
        (document.matches || []).some((name) => names[String(name).toLowerCase()] === true);

    return entries.filter((entry) =>
        isCurrentViewDocument(entry.documentId) ||
        entry.documents.some((document) => matchesCurrentView(document))
    );
};

/* ------------------------------------------------------------------ */
/* Service -> row mapping                                             */
/* ------------------------------------------------------------------ */
/**
 * Reads `record.a.b` and normalizes it to a trimmed string. Returns "" when a
 * step is missing, so the mapping never depends on one exact casing.
 */
const readAuditField = (record: any, path: string): string => {
    const parts = path.split(".");
    let value: any = record;

    for (let index = 0; index < parts.length; index++) {
        if (value === undefined || value === null) {
            return "";
        }
        value = value[parts[index]];
    }

    /* Plain values only - person columns are read through "Field.Title". */
    if (value === undefined || value === null || typeof value === "object") {
        return "";
    }

    return String(value).trim();
};

/** First non-empty value among the accepted spellings of one service field. */
const pickAuditField = (record: any, paths: string[]): string => {
    for (let index = 0; index < paths.length; index++) {
        const value = readAuditField(record, paths[index]);
        if (value !== "") {
            return value;
        }
    }
    return "";
};

const readAuditDate = (record: any): Date | null => {
    const raw = pickAuditField(record, ["Created", "created", "CreatedDate", "createdDate", "SentDate", "sentDate"]);
    if (raw === "") {
        return null;
    }

    const parsed = new Date(raw);
    return isNaN(parsed.getTime()) ? null : parsed;
};

/** "18 Nov 2026, 11:56 AM" - day, short month, year, 12-hour hour + minute. */
const formatAuditDate = (date: Date | null): string => {
    if (!date) {
        return "";
    }

    const formatted = date.toLocaleString("en-GB", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: true
    });

    /* en-GB lower-cases the meridiem; the design uses AM / PM. */
    return formatted.replace(/\s?(am|pm)\b/i, (match) => ` ${match.trim().toUpperCase()}`);
};

/**
 * Normalizes ONE item of "ProjectTransmittals". This is the single place where
 * service fields are bound to table columns:
 *
 *   ReferenceNo     -> TRANSMITTAL
 *   Created By      -> SENT BY
 *   Created         -> DATE & TIME SENT
 *   Name            -> DOCUMENTS
 *   PublisherEmail  -> RECIPIENT
 *   Remark          -> REMARK
 *
 * Both `ReferenceNo` / `referenceNo`, `CreatedBy` / `createdBy`,
 * `PublisherEmail` / `publisherEmail`, ... are accepted.
 */
export const mapAuditRow = (record: any): ITransmittalAuditLogEntry => {
    const created = readAuditDate(record);
    const referenceNo = pickAuditField(record, ["ReferenceNo", "referenceNo"]);
    /* Every spelling of the file name, in display priority order. */
    const documentNames: string[] = [];
    ["Name", "name", "ActualName", "actualName", "FileLeafRef", "fileLeafRef", "File.Name"].forEach((path) => {
        const value = readAuditField(record, path);
        if (value !== "" && documentNames.indexOf(value) === -1) {
            documentNames.push(value);
        }
    });
    const documentName = documentNames.length > 0 ? documentNames[0] : "";
    const itemId = pickAuditField(record, ["Id", "ID", "id"]);

    return {
        id: itemId !== "" ? `tal-${itemId}` : `tal-${referenceNo || documentName || "row"}`,
        transmittalNo: referenceNo,
        sentBy: pickAuditField(record, [
            "CreatedBy.Title",
            "CreatedBy",
            "createdBy",
            "Author.Title",
            "Author",
            "author"
        ]),
        role: "",
        dateTime: formatAuditDate(created),
        recipient: pickAuditField(record, ["PublisherEmail", "publisherEmail", "Recipient", "recipient"]),
        remark: pickAuditField(record, ["Remark", "remark", "Description", "description"]),
        documents: documentName === "" ? [] : [{ name: documentName, matches: documentNames }],
        folderPath: pickAuditField(record, ["FileDirRef", "fileDirRef"]),
        createdTime: created ? created.getTime() : 0
    };
};

/** Maps a service response to rows, newest first (not grouped yet). */
export const mapAuditRecords = (records: any[]): ITransmittalAuditLogEntry[] =>
    (records || [])
        .map(mapAuditRow)
        .sort((left, right) => (right.createdTime || 0) - (left.createdTime || 0));

/**
 * Folds every copy that shares a ReferenceNo into ONE row (one transmittal,
 * N documents). The input is expected newest first, so the newest copy wins the
 * header values; the file names are de-duplicated. Rows are cloned, never
 * mutated, so the caller can group the same records repeatedly.
 */
export const groupAuditRows = (rows: ITransmittalAuditLogEntry[]): ITransmittalAuditLogEntry[] => {
    const grouped: ITransmittalAuditLogEntry[] = [];
    const byTransmittal: { [key: string]: ITransmittalAuditLogEntry } = {};

    (rows || []).forEach((row) => {
        const key = (row.transmittalNo || row.id).toLowerCase();
        const existing = byTransmittal[key];

        if (!existing) {
            const copy: ITransmittalAuditLogEntry = { ...row, documents: row.documents.slice() };
            byTransmittal[key] = copy;
            grouped.push(copy);
            return;
        }

        /* Same transmittal: newest header values win, extra files are added. */
        row.documents.forEach((document) => {
            const alreadyListed = existing.documents.some(
                (current) => current.name.toLowerCase() === document.name.toLowerCase()
            );
            if (!alreadyListed) {
                existing.documents.push(document);
            }
        });
        if (!existing.remark && row.remark) {
            existing.remark = row.remark;
        }
        if (!existing.recipient && row.recipient) {
            existing.recipient = row.recipient;
        }
        if (!existing.sentBy && row.sentBy) {
            existing.sentBy = row.sentBy;
        }
        if (!existing.dateTime && row.dateTime) {
            existing.dateTime = row.dateTime;
        }
    });

    return grouped;
};

/** Map + group a whole service response (no folder scoping). */
export const mapAuditRows = (records: any[]): ITransmittalAuditLogEntry[] =>
    groupAuditRows(mapAuditRecords(records));

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */
const AUDIT_LOG_PANEL_ID = "transmittal-audit-log-panel";
const STYLE_ELEMENT_ID = "dms-transmittal-audit-log-styles";

/** "1" -> "0001" without relying on String.prototype.padStart (es5 target). */
const padSequence = (value: number): string => {
    const text = String(value);
    return text.length >= 4 ? text : `${"0000".substring(0, 4 - text.length)}${text}`;
};

/**
 * Next transmittal number for the given year, e.g. TR-2026-0002.
 * Numbers already present in the log are skipped, so an empty log restarts
 * the sequence at 0001 for that year.
 */
export const buildTransmittalNumber = (
    entries: ITransmittalAuditLogEntry[],
    date: Date = new Date()
): string => {
    const prefix = `TR-${date.getFullYear()}-`;
    let highest = 0;

    entries.forEach((entry) => {
        const value = entry.transmittalNo || "";
        if (value.indexOf(prefix) !== 0) {
            return;
        }
        const sequence = parseInt(value.substring(prefix.length), 10);
        if (!isNaN(sequence) && sequence > highest) {
            highest = sequence;
        }
    });

    return `${prefix}${padSequence(highest + 1)}`;
};

/** "1 document" / "2 documents". */
const formatDocumentCount = (count: number): string => `${count} document${count === 1 ? "" : "s"}`;

/* ------------------------------------------------------------------ */
/* Styles - injected once, scoped to the `tal-` prefix                 */
/* ------------------------------------------------------------------ */
const AUDIT_LOG_STYLES = `
.tal-toggle,
.tal-toggle *,
.tal-panel,
.tal-panel * {
  box-sizing: border-box;
}

/* ---------- Action bar toggle ---------- */
.tal-toggle {
  display: inline-flex;
  align-items: center;
  gap: 8px;
  height: 44px;
  padding: 0 12px 0 16px;
  border: 1px solid #9DBCF5;
  border-radius: 16px;
  background-color: #E8F0FE;
  color: #1A73E8;
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  font-size: 14.5px;
  font-weight: 600;
  line-height: 1;
  cursor: pointer;
  white-space: nowrap;
  box-shadow: 0 1px 2px rgba(26, 115, 232, 0.08);
  transition: background-color 160ms ease, border-color 160ms ease, box-shadow 160ms ease;
  /* Fade in smoothly when the button appears (location gains a document). */
  animation: tal-toggle-fade-in 150ms ease-out both;
}

@keyframes tal-toggle-fade-in {
  from {
    opacity: 0;
    transform: translateY(-2px);
  }
  to {
    opacity: 1;
    transform: translateY(0);
  }
}

.tal-toggle:hover {
  background-color: #DCE8FD;
  border-color: #7FA9F1;
  box-shadow: 0 4px 12px rgba(26, 115, 232, 0.16);
}

.tal-toggle--open,
.tal-toggle--open:hover {
  background-color: #D2E3FC;
  border-color: #7FA9F1;
  box-shadow: inset 0 1px 2px rgba(26, 115, 232, 0.16), 0 4px 12px rgba(26, 115, 232, 0.14);
}

.tal-toggle:focus-visible {
  outline: 2px solid #1A73E8;
  outline-offset: 2px;
}

.tal-toggle-icon {
  width: 18px;
  height: 18px;
  flex: 0 0 auto;
}

.tal-toggle-badge {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 22px;
  height: 22px;
  padding: 0 6px;
  border-radius: 999px;
  background-color: #1A73E8;
  color: #FFFFFF;
  font-size: 12px;
  font-weight: 600;
  line-height: 1;
}

/* ---------- Collapsible wrapper (smooth slide + fade) ---------- */
/*
 * The card lives in normal document flow between the action bar and the
 * documents table. The wrapper keeps its own box (a 0fr grid row) while
 * closed, so opening/closing only animates height + the 20px gap below the
 * card - the table simply moves down instead of being replaced.
 */
.tal-collapse {
  display: grid;
  grid-template-rows: 0fr;
  margin-top: 0;
  margin-bottom: 0;
  opacity: 0;
  visibility: hidden;
  pointer-events: none;
  transition: grid-template-rows 320ms cubic-bezier(0.4, 0, 0.2, 1),
              margin-bottom 320ms cubic-bezier(0.4, 0, 0.2, 1),
              opacity 200ms ease,
              visibility 0s linear 320ms;
}

.tal-collapse--open {
  grid-template-rows: 1fr;
  margin-top: 0;
  margin-bottom: 20px;
  opacity: 1;
  visibility: visible;
  pointer-events: auto;
  transition: grid-template-rows 320ms cubic-bezier(0.4, 0, 0.2, 1),
              margin-bottom 320ms cubic-bezier(0.4, 0, 0.2, 1),
              opacity 260ms ease 60ms,
              visibility 0s linear 0s;
}

.tal-collapse-inner {
  min-height: 0;
  overflow: hidden;
}

/* ---------- Panel ---------- */
.tal-panel {
  /*
   * Compact, capped card: the title/subtitle header stays pinned at the top
   * and only the table (or the empty state) scrolls inside. A long log
   * therefore never pushes the document table off the screen.
   */
  display: flex;
  flex-direction: column;
  max-height: 360px;
  overflow: hidden;
  background-color: #FFFFFF;
  border: 1px solid #E5E7EB;
  border-radius: 24px;
  box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04), 0 18px 44px rgba(15, 23, 42, 0.08);
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  font-size: 14px;
  color: #111827;
  transform: translateY(-6px);
  transition: transform 320ms cubic-bezier(0.4, 0, 0.2, 1);
}

.tal-collapse--open .tal-panel {
  transform: translateY(0);
}

.tal-panel-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  padding: 24px;
  flex: 0 0 auto;
  flex-wrap: wrap;
  background-color: #FFFFFF;
}

.tal-header-main {
  display: flex;
  align-items: flex-start;
  gap: 14px;
  min-width: 0;
}

.tal-header-icon {
  width: 26px;
  height: 26px;
  flex: 0 0 auto;
  margin-top: 3px;
  color: #1A73E8;
}

.tal-header-text {
  min-width: 0;
}

.tal-title {
  margin: 0;
  font-size: 24px;
  line-height: 1.3;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: #111827;
}

.tal-subtitle {
  margin: 6px 0 0;
  font-size: 14px;
  line-height: 1.55;
  color: #6B7280;
}

.tal-subtitle strong {
  color: #111827;
  font-weight: 600;
}

.tal-divider {
  height: 1px;
  flex: 0 0 auto;
  background-color: #E5E7EB;
}

/* ---------- Table ---------- */
.tal-table-scroll {
  width: 100%;
  flex: 1 1 auto;
  min-height: 0;
  overflow-x: auto;
  /* The log scrolls INSIDE the card instead of growing the page. */
  overflow-y: auto;
  -webkit-overflow-scrolling: touch;
  background-color: #FFFFFF;
}

.tal-table {
  width: 100%;
  min-width: 1040px;
  margin: 0;
  border-collapse: collapse;
  border-spacing: 0;
  background-color: #FFFFFF;
}

.tal-table thead th {
  padding: 14px 20px;
  background-color: #F9FAFB;
  border-bottom: 1px solid #E5E7EB;
  color: #6B7280;
  font-size: 14px;
  font-weight: 600;
  line-height: 1.2;
  letter-spacing: 0.06em;
  text-align: left;
  text-transform: uppercase;
  white-space: nowrap;
}

.tal-table tbody td {
  padding: 20px;
  border-bottom: 1px solid #E5E7EB;
  color: #374151;
  font-size: 16px;
  line-height: 1.45;
  vertical-align: middle;
  text-align: left;
}

.tal-table tbody tr {
  transition: background-color 150ms ease;
}

.tal-table tbody tr:hover {
  background-color: #F9FAFB;
}

.tal-table tbody tr:last-child td {
  border-bottom: 0;
}

.tal-cell-transmittal {
  color: #111827;
  font-weight: 700;
  white-space: nowrap;
}

.tal-person-name {
  color: #111827;
  font-weight: 500;
}

.tal-person-role {
  margin-top: 2px;
  color: #9CA3AF;
  font-size: 13px;
  line-height: 1.4;
}

.tal-date {
  color: #374151;
  white-space: nowrap;
}

.tal-recipient {
  color: #374151;
  word-break: break-word;
}

.tal-doc-count {
  color: #111827;
  font-weight: 500;
}

.tal-doc-name {
  margin-top: 2px;
  color: #9CA3AF;
  font-size: 13px;
  line-height: 1.4;
}

.tal-folder {
  min-width: 220px;
  color: #4B5563;
  font-size: 15px;
  line-height: 1.5;
  overflow-wrap: anywhere;
  word-break: break-word;
}

/* Remark: the narrow, wrapping last data column. */
.tal-remark {
  width: 160px;
  min-width: 160px;
  max-width: 200px;
  color: #374151;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.tal-cell-action {
  white-space: nowrap;
}

.tal-cover-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 28px;
  padding: 0 12px;
  border: 1px solid #9DBCF5;
  border-radius: 8px;
  background-color: #FFFFFF;
  color: #1A73E8;
  font-family: inherit;
  font-size: 12.5px;
  font-weight: 600;
  line-height: 1;
  white-space: nowrap;
  cursor: pointer;
  transition: background-color 150ms ease, border-color 150ms ease;
}

.tal-cover-btn:hover {
  background-color: #E8F0FE;
  border-color: #6FA0EE;
}

.tal-cover-btn:focus-visible {
  outline: 2px solid #1A73E8;
  outline-offset: 2px;
}

/* ---------- Loading skeleton ---------- */
.tal-skeleton-bar {
  display: block;
  height: 12px;
  border-radius: 999px;
  background-color: #E5E7EB;
  animation: tal-skeleton-pulse 1.2s ease-in-out infinite;
}

@keyframes tal-skeleton-pulse {
  0%,
  100% {
    opacity: 1;
  }
  50% {
    opacity: 0.45;
  }
}

/* ---------- Error ---------- */
.tal-error {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  flex: 1 1 auto;
  min-height: 0;
  padding: 24px;
  overflow-y: auto;
  background-color: #FFFFFF;
}

.tal-error-title {
  margin: 0;
  color: #111827;
  font-size: 15px;
  font-weight: 600;
}

.tal-retry-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  height: 34px;
  padding: 0 16px;
  border: 1px solid #9DBCF5;
  border-radius: 10px;
  background-color: #FFFFFF;
  color: #1A73E8;
  font-family: inherit;
  font-size: 13px;
  font-weight: 600;
  line-height: 1;
  white-space: nowrap;
  cursor: pointer;
  flex: 0 0 auto;
  transition: background-color 150ms ease, border-color 150ms ease;
}

.tal-retry-btn:hover {
  background-color: #E8F0FE;
  border-color: #6FA0EE;
}

.tal-retry-btn:focus-visible {
  outline: 2px solid #1A73E8;
  outline-offset: 2px;
}

/* ---------- Empty state ---------- */
.tal-empty {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 10px;
  flex: 1 1 auto;
  min-height: 0;
  padding: 32px 24px;
  overflow-y: auto;
  background-color: #FFFFFF;
  text-align: center;
}

.tal-empty-icon {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 48px;
  height: 48px;
  border-radius: 16px;
  background-color: #E8F0FE;
  color: #1A73E8;
}

.tal-empty-icon svg {
  width: 24px;
  height: 24px;
}

.tal-empty-title {
  margin: 0;
  color: #111827;
  font-size: 16px;
  font-weight: 600;
}

.tal-empty-text {
  max-width: 420px;
  margin: 0;
  color: #6B7280;
  font-size: 14px;
  line-height: 1.55;
}

.tal-visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
  border: 0;
}

/* ---------- Small screens ---------- */
@media (max-width: 900px) {
  .tal-panel {
    border-radius: 20px;
  }

  .tal-panel-header {
    padding: 18px;
  }

  .tal-title {
    font-size: 20px;
  }

  .tal-table {
    min-width: 880px;
  }

  .tal-table thead th {
    padding: 12px 14px;
    font-size: 12px;
  }

  .tal-table tbody td {
    padding: 14px;
    font-size: 14px;
  }

  .tal-folder {
    min-width: 180px;
    font-size: 13.5px;
  }
}

/* ---------- Motion preferences ---------- */
@media (prefers-reduced-motion: reduce) {
  .tal-toggle,
  .tal-collapse,
  .tal-collapse--open,
  .tal-panel,
  .tal-table tbody tr,
  .tal-cover-btn,
  .tal-retry-btn,
  .tal-skeleton-bar {
    transition: none;
    animation: none;
  }
}
`;

/** Adds the scoped stylesheet to <head> exactly once. */
const useAuditLogStyles = (): void => {
    useEffect(() => {
        if (typeof document === "undefined" || document.getElementById(STYLE_ELEMENT_ID)) {
            return;
        }

        const styleElement = document.createElement("style");
        styleElement.id = STYLE_ELEMENT_ID;
        styleElement.textContent = AUDIT_LOG_STYLES;
        document.head.appendChild(styleElement);
    }, []);
};

/* ------------------------------------------------------------------ */
/* 1. Action bar toggle                                                */
/* ------------------------------------------------------------------ */
export const TransmittalAuditLogToggle: React.FC<ITransmittalAuditLogToggleProps> = ({
    isOpen,
    count,
    onToggle
}) => {
    useAuditLogStyles();

    return (
        <button
            type="button"
            className={`tal-toggle${isOpen ? " tal-toggle--open" : ""}`}
            onClick={onToggle}
            aria-expanded={isOpen}
            aria-controls={AUDIT_LOG_PANEL_ID}
            aria-label={`Audit Log, ${count} ${count === 1 ? "entry" : "entries"}`}
            title={isOpen ? "Hide the transmittal audit log" : "Show the transmittal audit log"}
            data-testid="button-audit-log"
        >
            <FileTextIcon className="tal-toggle-icon" />
            <span className="tal-toggle-label">Audit Log</span>
            <span className="tal-toggle-badge" aria-hidden="true">{count}</span>
        </button>
    );
};

/* ------------------------------------------------------------------ */
/* 2. Collapsible panel                                                */
/* ------------------------------------------------------------------ */
export const TransmittalAuditLog: React.FC<ITransmittalAuditLogProps> = ({
    isOpen,
    context,
    documents,
    folderPath,
    refreshToken,
    onCountChange,
    onViewCover,
    onRequestClose
}) => {
    useAuditLogStyles();

    /* Copies read from "ProjectTransmittals" (newest first, one per document). */
    const [records, setRecords] = useState<ITransmittalAuditLogEntry[]>([]);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [error, setError] = useState<string>("");
    /* "Retry" bumps this so the very same service call runs again. */
    const [retryToken, setRetryToken] = useState<number>(0);

    /*
     * One fetch per folder (plus one after a send or a retry). Deliberately NOT
     * tied to `isOpen`, so opening and closing the card never hits the service
     * again.
     */
    useEffect(() => {
        let isCancelled = false;

        const loadAuditLog = async (): Promise<void> => {
            if (!context) {
                setRecords([]);
                setIsLoading(false);
                return;
            }

            setIsLoading(true);
            setError("");

            try {
                const items = await getTransmittalAuditLog(context);
                if (!isCancelled) {
                    setRecords(mapAuditRecords(items));
                }
            } catch (loadError) {
                if (!isCancelled) {
                    setRecords([]);
                    setError(
                        loadError instanceof Error && loadError.message
                            ? loadError.message
                            : "Could not load audit log"
                    );
                }
            } finally {
                if (!isCancelled) {
                    setIsLoading(false);
                }
            }
        };

        loadAuditLog();

        return () => {
            isCancelled = true;
        };
    }, [context, folderPath, refreshToken, retryToken]);

    /*
     * Scope the log to the open folder, then fold the folder's copies that share
     * a ReferenceNo into one row. The copies in "ProjectTransmittals" do not
     * store the folder they were sent from, so the existing helper matches them
     * against the file names currently on screen.
     */
    const entries = useMemo(
        () => groupAuditRows(filterAuditLogEntries(records, documents || [])),
        [records, documents]
    );

    /* Blue badge: distinct ReferenceNo values currently on screen. */
    useEffect(() => {
        if (!onCountChange) {
            return;
        }

        const seen: { [key: string]: boolean } = {};
        entries.forEach((entry) => {
            seen[(entry.transmittalNo || entry.id).toLowerCase()] = true;
        });
        onCountChange(Object.keys(seen).length);
    }, [entries, onCountChange]);

    const retryAuditLog = useCallback(() => {
        setRetryToken((value) => value + 1);
    }, []);

    /* Optional Escape-to-close, wired by the page that owns the open state. */
    useEffect(() => {
        if (!isOpen || !onRequestClose) {
            return undefined;
        }

        const handleKeyDown = (event: KeyboardEvent): void => {
            if (event.key === "Escape") {
                onRequestClose();
            }
        };

        document.addEventListener("keydown", handleKeyDown);
        return () => document.removeEventListener("keydown", handleKeyDown);
    }, [isOpen, onRequestClose]);

    const latestEntry = entries.length > 0 ? entries[0] : null;

    return (
        <section
            id={AUDIT_LOG_PANEL_ID}
            className={`tal-collapse${isOpen ? " tal-collapse--open" : ""}`}
            aria-hidden={!isOpen}
            data-testid="panel-transmittal-audit-log"
        >
            <div className="tal-collapse-inner">
                <div className="tal-panel" role="region" aria-label="Transmittal audit log">
                    <div className="tal-panel-header">
                        <div className="tal-header-main">
                            <FileTextIcon className="tal-header-icon" />
                            <div className="tal-header-text">
                                <h2 className="tal-title">Transmittal Audit Log</h2>
                                {latestEntry ? (
                                    <p className="tal-subtitle">
                                        Last transmittal <strong>{latestEntry.transmittalNo}</strong> sent by{" "}
                                        <strong>{latestEntry.sentBy}</strong> on <strong>{latestEntry.dateTime}</strong>.
                                    </p>
                                ) : isLoading ? (
                                    <p className="tal-subtitle">Loading transmittal audit log&hellip;</p>
                                ) : (
                                    <p className="tal-subtitle">
                                        No transmittals have been sent from this folder yet.
                                    </p>
                                )}
                            </div>
                        </div>
                    </div>

                    <div className="tal-divider" />

                    {error !== "" && !isLoading ? (
                        <div className="tal-error" role="alert" title={error}>
                            <p className="tal-error-title">Could not load audit log</p>
                            <button
                                type="button"
                                className="tal-retry-btn"
                                onClick={retryAuditLog}
                                data-testid="button-retry-audit-log"
                            >
                                Retry
                            </button>
                        </div>
                    ) : isLoading || entries.length > 0 ? (
                        <div className="tal-table-scroll">
                            <table className="tal-table">
                                <caption className="tal-visually-hidden">Transmittal audit log entries</caption>
                                <thead>
                                    <tr>
                                        <th scope="col">Transmittal</th>
                                        <th scope="col">Sent by</th>
                                        <th scope="col">Date &amp; time sent</th>
                                        <th scope="col">Recipient</th>
                                        <th scope="col">Documents</th>
                                        <th scope="col">Remark</th>
                                        <th scope="col">
                                            <span className="tal-visually-hidden">Actions</span>
                                        </th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {isLoading ? (
                                        [0, 1, 2].map((skeletonRow) => (
                                            <tr key={`tal-skeleton-${skeletonRow}`} aria-hidden="true">
                                                <td><span className="tal-skeleton-bar" style={{ width: "70%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "60%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "82%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "74%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "56%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "88%" }} /></td>
                                                <td><span className="tal-skeleton-bar" style={{ width: "46px" }} /></td>
                                            </tr>
                                        ))
                                    ) : (
                                        entries.map((entry) => (
                                            <tr key={entry.id}>
                                                <td className="tal-cell-transmittal">{entry.transmittalNo || "-"}</td>
                                                <td>
                                                    <div className="tal-person-name">{entry.sentBy || "-"}</div>
                                                </td>
                                                <td className="tal-date">{entry.dateTime || "-"}</td>
                                                <td className="tal-recipient">{entry.recipient || "-"}</td>
                                                <td>
                                                    <div className="tal-doc-count">
                                                        {formatDocumentCount(entry.documents.length)}
                                                    </div>
                                                    <div className="tal-doc-name">
                                                        {entry.documents.map((document, index) => (
                                                            <div key={`${document.name}-${index}`}>{document.name}</div>
                                                        ))}
                                                    </div>
                                                </td>
                                                <td className="tal-remark">{entry.remark || "-"}</td>
                                                <td className="tal-cell-action">
                                                    <button
                                                        type="button"
                                                        className="tal-cover-btn"
                                                        onClick={() => onViewCover?.(entry)}
                                                        title={`View the cover sheet for ${entry.transmittalNo}`}
                                                    >
                                                        Cover
                                                    </button>
                                                </td>
                                            </tr>
                                        ))
                                    )}
                                </tbody>
                            </table>
                        </div>
                    ) : (
                        <div className="tal-empty">
                            <div className="tal-empty-icon">
                                <FileTextIcon />
                            </div>
                            <p className="tal-empty-title" data-testid="text-audit-log-empty">
                                No transmittals sent yet
                            </p>
                            <p className="tal-empty-text">
                                Every transmittal you send from this folder will be recorded here with its
                                recipient, documents and folder path.
                            </p>
                        </div>
                    )}
                </div>
            </div>
        </section>
    );
};

export default React.memo(TransmittalAuditLog);
