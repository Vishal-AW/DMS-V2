import * as React from "react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as ReactDOM from "react-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Dismiss20Regular, Document20Regular } from "@fluentui/react-icons";
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import {
    formatTransmittalVersion,
    ISendTransmittalForm,
    ITransmittalDocument,
    ITransmittalDocumentResult,
    toTransmittalDocument,
    TRANSMITTAL_LIBRARY_NAME,
    uploadTransmittalDocuments
} from "../../../../Services/TransmittalService";
import "./SendTransmittalModal.css";

export interface ISendTransmittalModalProps {
    context: WebPartContext;
    /** Ticked table rows (ReusableDataTable row shape, passed through unchanged). */
    documents: any[];
    /** Current folder path used to prefill the subject, e.g. "Projects / 2026 / Tower A". */
    folderPath?: string;
    onDismiss: () => void;
    /**
     * Called once every copy was added to "ProjectTransmittals" with its metadata.
     * Hook the cover-sheet / mailto step here; it is never called for a failed run.
     * It may return a promise; the popup stays in "Sending…" until it resolves and
     * only closes afterwards.
     */
    onSent?: (results: ITransmittalDocumentResult[], form: ISendTransmittalForm) => void | Promise<void>;
}

interface IFileBadge {
    /** Badge text ("W", "X", "PDF"). Empty means "render the generic document icon". */
    label: string;
    modifierClass: string;
}

const FOCUSABLE_SELECTOR = [
    "button:not([disabled])",
    "[href]",
    "input:not([disabled])",
    "select:not([disabled])",
    "textarea:not([disabled])",
    "[tabindex]:not([tabindex='-1'])"
].join(",");

const EMAIL_PATTERN = /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/;

/** Multiple addresses may be separated with a comma or a semicolon. */
const parseEmailList = (value: string): string[] =>
    value.split(/[,;]+/).map((entry) => entry.trim()).filter((entry) => entry.length > 0);

const validateClientEmail = (value: string): string => {
    const emails = parseEmailList(value);
    if (emails.length === 0) {
        return "Enter at least one client email address.";
    }

    const invalid = emails.filter((email) => !EMAIL_PATTERN.test(email));
    if (invalid.length > 0) {
        return invalid.length === 1
            ? "Enter a valid email address, separated by a comma or semicolon."
            : "Enter valid email addresses, separated by commas or semicolons.";
    }

    return "";
};

const getFileBadge = (document: ITransmittalDocument): IFileBadge => {
    const extension = document.name.indexOf(".") > -1
        ? document.name.split(".").pop()!.toLowerCase()
        : "";

    if (extension === "doc" || extension === "docx") {
        return { label: "W", modifierClass: "stm-doc-badge-word" };
    }
    if (extension === "xls" || extension === "xlsx" || extension === "csv") {
        return { label: "X", modifierClass: "stm-doc-badge-excel" };
    }
    if (extension === "pdf") {
        return { label: "PDF", modifierClass: "stm-doc-badge-pdf" };
    }

    return { label: "", modifierClass: "stm-doc-badge-other" };
};

/** "2026-00120 · v0.4 · Published" - empty parts are skipped. */
const getDocumentMeta = (document: ITransmittalDocument): string => {
    const parts = [
        document.referenceNo || "",
        formatTransmittalVersion(document.version),
        document.status || ""
    ].filter((part) => part.length > 0);

    return parts.join(" · ");
};



const SendTransmittalModal: React.FC<ISendTransmittalModalProps> = ({
    context,
    documents,
    folderPath,
    onDismiss,
    onSent
}) => {
    const [clientEmail, setClientEmail] = useState("");
    /**
     * E-mail list of the picked users, kept so the PeoplePicker can be re-hydrated
     * (same pattern as the TileAdmin picker in TileForm).
     */
    const [selectedClientEmails, setSelectedClientEmails] = useState<string[]>([]);
    const [subject, setSubject] = useState<string>((): string => {
        const path = (folderPath || "").trim();
        return path ? `Transmittal · ${path}` : "Transmittal";
    });
    const [remarks, setRemarks] = useState("");
    const [emailError, setEmailError] = useState("");
    const [submitError, setSubmitError] = useState("");
    const [isSending, setIsSending] = useState(false);
    /** File names of a completed run; as soon as it is set the success view is shown. */
    const [sentFileNames, setSentFileNames] = useState<string[] | null>(null);

    /** Wrapper around the PeoplePicker - used to focus the picker's inner input. */
    const pickerWrapRef = useRef<HTMLDivElement>(null);
    const cardRef = useRef<HTMLDivElement>(null);
    const doneButtonRef = useRef<HTMLButtonElement>(null);
    const previouslyFocusedRef = useRef<HTMLElement | null>(null);
    /**
     * Results of a successful upload. Kept so a failure in the cover-sheet / e-mail
     * step can be retried without copying the documents a second time.
     */
    const uploadedResultsRef = useRef<ITransmittalDocumentResult[] | null>(null);
    /** True while the success view is shown (see the reset effect below). */
    const successVisibleRef = useRef(false);

    const transmittalDocuments = useMemo<ITransmittalDocument[]>(
        () => (documents || []).map(toTransmittalDocument),
        [documents]
    );

    // A fresh popup (or a different selection) always starts with a clean upload state.
    // The success view is the exception: the parent clears the table selection right
    // after a successful send, which also changes `documents`, and that must not hide
    // the confirmation while it is on screen.
    useEffect(() => {
        uploadedResultsRef.current = null;
        if (!successVisibleRef.current) {
            setSentFileNames(null);
        }
    }, [transmittalDocuments]);

    /**
     * Writes to uploadedResultsRef go through this helper: a direct `.current =` after
     * an `await` trips `require-atomic-updates`, although the ref is exactly what has
     * to survive the retry.
     */
    const storeUploadedResults = useCallback((value: ITransmittalDocumentResult[] | null): void => {
        uploadedResultsRef.current = value;
    }, []);

    const peoplePickerContext = React.useMemo<IPeoplePickerContext>(() => ({
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any
    }), [context]);

    const focusPicker = useCallback(() => {
        const wrapper = pickerWrapRef.current;
        if (!wrapper) {
            return;
        }

        const input = wrapper.querySelector<HTMLInputElement>(".ms-BasePicker-input")
            || wrapper.querySelector<HTMLInputElement>("input");
        input?.focus();
    }, []);

    // Focus the first field, lock background scrolling, and give focus back on close.
    useEffect(() => {
        previouslyFocusedRef.current = document.activeElement as HTMLElement;
        const previousOverflow = document.body.style.overflow;
        document.body.style.overflow = "hidden";

        const focusTimer = window.setTimeout(() => {
            focusPicker();
        }, 0);

        return () => {
            window.clearTimeout(focusTimer);
            document.body.style.overflow = previousOverflow;
            const previous = previouslyFocusedRef.current;
            if (previous && typeof previous.focus === "function") {
                previous.focus();
            }
        };
    }, [focusPicker]);

    // Move the focus to "Done" as soon as the success view replaces the form.
    useEffect(() => {
        if (sentFileNames) {
            doneButtonRef.current?.focus();
        }
    }, [sentFileNames]);

    /**
     * Leaves the popup. The success flag is cleared here too, so a popup that stays
     * mounted (or is opened again) never re-opens on a stale confirmation.
     */
    const dismiss = useCallback(() => {
        successVisibleRef.current = false;
        setSentFileNames(null);
        onDismiss();
    }, [onDismiss]);

    const handleOverlayMouseDown = useCallback((event: React.MouseEvent<HTMLDivElement>) => {
        if (event.target !== event.currentTarget || isSending) {
            return;
        }
        dismiss();
    }, [dismiss, isSending]);

    const handleCardKeyDown = useCallback((event: React.KeyboardEvent<HTMLDivElement>) => {
        if (event.key === "Escape") {
            event.stopPropagation();
            if (!isSending) {
                dismiss();
            }
            return;
        }

        if (event.key !== "Tab") {
            return;
        }

        const card = cardRef.current;
        if (!card) {
            return;
        }

        const focusable = Array.prototype.slice.call(card.querySelectorAll(FOCUSABLE_SELECTOR)) as HTMLElement[];
        const visible = focusable.filter((element) => element.offsetParent !== null || element === document.activeElement);
        if (visible.length === 0) {
            event.preventDefault();
            return;
        }

        const first = visible[0];
        const last = visible[visible.length - 1];

        if (event.shiftKey && document.activeElement === first) {
            event.preventDefault();
            last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
            event.preventDefault();
            first.focus();
        }
    }, [dismiss, isSending]);

    const handleCancel = useCallback(() => {
        if (!isSending) {
            dismiss();
        }
    }, [dismiss, isSending]);

    const handleSubmit = useCallback(async () => {
        if (isSending) {
            return;
        }

        if (!clientEmail.trim()) {
            setEmailError("Client email is required.");
            focusPicker();
            return;
        }

        const validationMessage = validateClientEmail(clientEmail);
        if (validationMessage) {
            setEmailError(validationMessage);
            focusPicker();
            return;
        }

        setEmailError("");
        setSubmitError("");
        setIsSending(true);

        // Step 1 - copy the documents. Skipped when an earlier attempt already did it,
        // so retrying the cover-sheet step can never create duplicate copies.
        let results = uploadedResultsRef.current;
        if (!results) {
            try {
                results = await uploadTransmittalDocuments(context, {
                    documents: transmittalDocuments,
                    remarks
                });
                storeUploadedResults(results);
            } catch (error) {
                setSubmitError(
                    error instanceof Error && error.message
                        ? error.message
                        : "The transmittal could not be sent. Please try again."
                );
                setIsSending(false);
                return;
            }
        }

        // Step 2 - cover sheet / e-mail. The copies already exist, so a failure here is
        // reported on its own and the upload is never repeated.
        try {
            await onSent?.(results, {
                clientEmails: parseEmailList(clientEmail),
                subject,
                remarks
            });
        } catch (error) {
            setSubmitError(
                `The documents were added to ${TRANSMITTAL_LIBRARY_NAME}, but the next step failed: ` +
                `${error instanceof Error && error.message ? error.message : "unknown error"}. ` +
                "Click Generate & Send again to retry only that step."
            );
            setIsSending(false);
            return;
        }

        // Success - the dialog stays open and shows the confirmation instead of closing.
        storeUploadedResults(null);
        successVisibleRef.current = true;
        setSentFileNames(results.map((result) => result.document.actualName || result.document.name));
        setIsSending(false);
    }, [clientEmail, context, focusPicker, isSending, onSent, remarks, storeUploadedResults, subject, transmittalDocuments]);

    const documentCount = transmittalDocuments.length;

    return ReactDOM.createPortal(
        <div className="stm-overlay" onMouseDown={handleOverlayMouseDown}>
            <div
                ref={cardRef}
                className="stm-card"
                role="dialog"
                aria-modal="true"
                aria-labelledby="stm-title"
                aria-describedby="stm-subtitle"
                onKeyDown={handleCardKeyDown}
            >
                <div className="stm-header">
                    <div className="stm-header-text">
                        <h2 id="stm-title" className="stm-title">Send Transmittal</h2>
                        <p id="stm-subtitle" className="stm-subtitle">
                            {sentFileNames
                                ? "Transmittal complete"
                                : `${documentCount} document${documentCount === 1 ? "" : "s"} selected`}
                        </p>
                    </div>
                    <button
                        type="button"
                        className="stm-close"
                        aria-label="Close send transmittal dialog"
                        onClick={handleCancel}
                        disabled={isSending}
                    >
                        <Dismiss20Regular />
                    </button>
                </div>

                {/* Form view - replaced by the success view once the copies are in. */}
                {!sentFileNames && (
                    <>
                        <div className="stm-field">
                            <label className="stm-label">
                                Client email <span aria-hidden="true" style={{ color: "#dc2626" }}>*</span>
                            </label>
                            <div
                                ref={pickerWrapRef}
                                className={`stm-picker${emailError ? " stm-picker-invalid" : ""}`}
                                aria-describedby={emailError ? "stm-client-email-error" : undefined}
                            >
                                <PeoplePicker
                                    context={peoplePickerContext}
                                    personSelectionLimit={5}
                                    showtooltip={true}
                                    showHiddenInUI={false}
                                    ensureUser={true}
                                    principalTypes={[PrincipalType.User]}
                                    placeholder="client@company.com"
                                    disabled={isSending}
                                    onChange={(users: any[]) => {
                                        // The picker exposes the user's e-mail in secondaryText.
                                        const emails = (users || [])
                                            .map((user) => (user && typeof user.secondaryText === "string"
                                                ? user.secondaryText.trim()
                                                : ""))
                                            .filter((email) => email.length > 0);

                                        setSelectedClientEmails(emails);
                                        setClientEmail(emails.join(", "));

                                        if (emails.length > 0) {
                                            setEmailError("");
                                        }
                                    }}
                                    defaultSelectedUsers={selectedClientEmails}
                                />
                            </div>
                            {emailError && (
                                <div id="stm-client-email-error" className="stm-field-error" role="alert">
                                    {emailError}
                                </div>
                            )}
                        </div>

                        <div className="stm-field">
                            <label className="stm-label" htmlFor="stm-subject">Subject</label>
                            <input
                                id="stm-subject"
                                className="stm-input"
                                type="text"
                                value={subject}
                                onChange={(event) => setSubject(event.target.value)}
                                disabled={isSending}
                            />
                        </div>

                        <div className="stm-field">
                            <label className="stm-label" htmlFor="stm-remarks">Remarks (optional)</label>
                            <textarea
                                id="stm-remarks"
                                className="stm-textarea"
                                rows={3}
                                value={remarks}
                                onChange={(event) => setRemarks(event.target.value)}
                                disabled={isSending}
                            />
                        </div>

                        <div className="stm-docs">
                            <div className="stm-docs-header">Documents included ({documentCount})</div>
                            <div className="stm-docs-list">
                                {documentCount === 0 && (
                                    <div className="stm-docs-empty">No documents selected.</div>
                                )}
                                {transmittalDocuments.map((document, index) => {
                                    const badge = getFileBadge(document);
                                    const meta = getDocumentMeta(document);

                                    return (
                                        <div className="stm-doc-row" key={`${document.id}-${index}`}>
                                            <span className={`stm-doc-badge ${badge.modifierClass}`} aria-hidden="true">
                                                {badge.label || <Document20Regular />}
                                            </span>
                                            <span className="stm-doc-name" title={document.actualName}>
                                                {document.actualName}
                                            </span>
                                            {meta && <span className="stm-doc-meta">{meta}</span>}
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                        <p className="stm-note">
                            Each selected document is copied into Project Transmittals and tagged with its reference, level and
                            the transmittal status. The original files stay untouched in the current folder.
                        </p>

                        {submitError && (
                            <div className="stm-global-error" role="alert">{submitError}</div>
                        )}

                        <div className="stm-actions">
                            <button
                                type="button"
                                className="stm-btn stm-btn-secondary"
                                onClick={handleCancel}
                                disabled={isSending}
                            >
                                Cancel
                            </button>
                            <button
                                type="button"
                                className="stm-btn stm-btn-primary"
                                onClick={handleSubmit}
                                disabled={isSending}
                                aria-busy={isSending}
                            >
                                {isSending ? "Sending…" : "Generate & Send"}
                            </button>
                        </div>
                    </>
                )}

                {sentFileNames && (
                    <div
                        className="stm-success"
                        role="status"
                        aria-live="polite"
                        data-testid="modal-transmittal-success"
                    >
                        <span className="stm-success-icon" aria-hidden="true">
                            <svg viewBox="0 0 24 24" width="34" height="34" focusable="false">
                                <path
                                    d="M5 12.6l4.6 4.6L19 7.6"
                                    fill="none"
                                    stroke="#15803d"
                                    strokeWidth="2.4"
                                    strokeLinecap="round"
                                    strokeLinejoin="round"
                                />
                            </svg>
                        </span>
                        <h3 className="stm-success-title">Transmittal sent</h3>
                        <p className="stm-success-text">
                            {`${sentFileNames.length} document${sentFileNames.length === 1 ? " was" : "s were"} added to Project Transmittals.`}
                        </p>
                        {sentFileNames.length > 0 && (
                            <ul className="stm-success-list">
                                {sentFileNames.map((fileName, index) => (
                                    <li className="stm-success-item" key={`${fileName}-${index}`} title={fileName}>
                                        {fileName}
                                    </li>
                                ))}
                            </ul>
                        )}
                        <div className="stm-actions">
                            <button
                                ref={doneButtonRef}
                                type="button"
                                className="stm-btn stm-btn-primary"
                                onClick={dismiss}
                            >
                                Done
                            </button>
                        </div>
                    </div>
                )}

            </div>
        </div>,
        document.body
    );
};

export default React.memo(SendTransmittalModal);
