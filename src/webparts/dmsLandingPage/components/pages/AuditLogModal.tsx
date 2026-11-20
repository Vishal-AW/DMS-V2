/* eslint-disable */
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { ITransmittalAuditLogEntry } from "./TransmittalAuditLog";

/**
 * AuditLogModal
 * ---------------------------------------------------------------------------
 * Popup (modal) presentation of the transmittal audit log.
 *
 * The workspace action bar keeps its "Audit Log" toggle; this component is the
 * overlay that opens on top of the page instead of replacing the document
 * table. It is rendered into `document.body` through a React portal, so the
 * grid behind it stays mounted with its selection, filters, sorting, and
 * scroll position completely untouched while the popup is open.
 *
 * Behaviour
 *  - Fixed, full-screen backdrop at z-index 50 (semi-transparent + blur).
 *  - Fade/scale in on open and fade/scale out on close.
 *  - Closes on the X button, a backdrop click, or the Escape key.
 *  - Locks `document.body` scroll while open and restores it on close.
 *  - Moves focus into the dialog, traps Tab inside it, and returns focus to
 *    the "Audit Log" toggle when it closes.
 *
 * The solution does not ship Tailwind CSS or lucide-react (see package.json),
 * so the design tokens below are hand-written, scoped styles (`alm-` prefix,
 * injected once) and the two icons are inline SVGs using lucide's path data.
 */

export interface IAuditLogModalProps {
    /** Whether the popup is open. Owned by the page. */
    isOpen: boolean;
    /** Entries for the current document view, newest first. */
    entries: ITransmittalAuditLogEntry[];
    /** Called by the X button, the backdrop and the Escape key. */
    onRequestClose: () => void;
    /** Optional "Cover" action for a row. */
    onViewCover?: (entry: ITransmittalAuditLogEntry) => void;
}

/** Matches the CSS exit animation (`alm-overlay-out` / `alm-dialog-out`). */
const EXIT_ANIMATION_MS = 160;
const STYLE_ELEMENT_ID = "dms-audit-log-modal-styles";
const TITLE_ID = "dms-audit-log-modal-title";

/** Elements that can receive focus while the dialog is open. */
const FOCUSABLE_SELECTOR = [
    "a[href]",
    "button:not([disabled])",
    "textarea:not([disabled])",
    "input:not([disabled]):not([type=\"hidden\"])",
    "select:not([disabled])",
    "[tabindex]:not([tabindex=\"-1\"])"
].join(",");

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

/** lucide-react `X`. */
const XIcon: React.FC<ILucideIconProps> = (props) => (
    <LucideIcon {...props}>
        <path d="M18 6 6 18" />
        <path d="m6 6 12 12" />
    </LucideIcon>
);

/** "1 document" / "2 documents". */
const formatDocumentCount = (count: number): string => `${count} document${count === 1 ? "" : "s"}`;

/* ------------------------------------------------------------------ */
/* Styles - injected once, scoped to the `alm-` prefix                 */
/* ------------------------------------------------------------------ */
const MODAL_STYLES_PART_ONE = `
.alm-overlay,
.alm-overlay *,
.alm-overlay *::before,
.alm-overlay *::after {
  box-sizing: border-box;
}

/* ---------- Backdrop ---------- */
.alm-overlay {
  position: fixed;
  inset: 0;
  /* Must stay above the page chrome (app header z-101/102, top nav z-1000)
     and below Fluent UI's Layer stacking context (z-index: 1000000), exactly
     like the Send Transmittal modal. */
  z-index: 999999;
  display: flex;
  align-items: flex-start;
  justify-content: center;
  padding: 24px;
  overflow-y: auto;
  background-color: rgba(15, 23, 42, 0.4);
  -webkit-backdrop-filter: blur(4px);
  backdrop-filter: blur(4px);
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif;
  color: #111827;
  overscroll-behavior: contain;
  animation: alm-overlay-in 180ms ease-out both;
}

.alm-overlay--closing {
  animation: alm-overlay-out 160ms ease-in both;
}

@keyframes alm-overlay-in {
  from { opacity: 0; }
  to { opacity: 1; }
}

@keyframes alm-overlay-out {
  from { opacity: 1; }
  to { opacity: 0; }
}

/* ---------- Dialog container ---------- */
.alm-dialog {
  position: relative;
  display: flex;
  flex-direction: column;
  width: 100%;
  max-width: 1152px;
  max-height: 85vh;
  background-color: #ffffff;
  border: 1px solid #E5E7EB;
  border-radius: 24px;
  box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.35), 0 12px 24px rgba(15, 23, 42, 0.12);
  overflow: hidden;
  outline: none;
  animation: alm-dialog-in 200ms cubic-bezier(0.16, 1, 0.3, 1) both;
}

.alm-dialog--closing {
  animation: alm-dialog-out 160ms ease-in both;
}

@keyframes alm-dialog-in {
  from { opacity: 0; transform: scale(0.97) translateY(6px); }
  to { opacity: 1; transform: scale(1) translateY(0); }
}

@keyframes alm-dialog-out {
  from { opacity: 1; transform: scale(1) translateY(0); }
  to { opacity: 0; transform: scale(0.98) translateY(4px); }
}

/* ---------- Header ---------- */
.alm-header {
  display: flex;
  align-items: flex-start;
  justify-content: space-between;
  gap: 16px;
  padding: 24px;
  background-color: #ffffff;
  border-bottom: 1px solid #E5E7EB;
}

.alm-header-main {
  display: flex;
  align-items: flex-start;
  gap: 14px;
  min-width: 0;
}

.alm-header-icon {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 44px;
  height: 44px;
  flex: 0 0 auto;
  border-radius: 12px;
  background-color: #E8F0FE;
  color: #1A73E8;
}

.alm-header-icon svg {
  width: 22px;
  height: 22px;
}

.alm-header-text {
  min-width: 0;
}

.alm-title {
  margin: 0;
  font-size: 24px;
  line-height: 1.25;
  font-weight: 600;
  letter-spacing: -0.01em;
  color: #111827;
}

.alm-subtitle {
  margin: 4px 0 0;
  font-size: 14px;
  line-height: 1.5;
  color: #6B7280;
}

.alm-subtitle strong {
  color: #111827;
  font-weight: 600;
}

.alm-close {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 40px;
  height: 40px;
  flex: 0 0 auto;
  padding: 0;
  border: 0;
  border-radius: 999px;
  background-color: transparent;
  color: #6B7280;
  font-family: inherit;
  cursor: pointer;
  transition: background-color 150ms ease, color 150ms ease;
}

.alm-close:hover {
  background-color: #F3F4F6;
  color: #111827;
}

.alm-close:focus-visible {
  outline: 2px solid #1A73E8;
  outline-offset: 2px;
}

.alm-close svg {
  width: 20px;
  height: 20px;
}
`;

const MODAL_STYLES_PART_TWO = `
/* ---------- Scrollable body ---------- */
.alm-body {
  flex: 1 1 auto;
  min-height: 0;
  overflow: auto;
  overscroll-behavior: contain;
  background-color: #ffffff;
}

/* ---------- Audit table ---------- */
.alm-table {
  width: 100%;
  min-width: 1060px;
  border-collapse: separate;
  border-spacing: 0;
  font-size: 14px;
}

.alm-table thead th {
  position: sticky;
  top: 0;
  z-index: 2;
  padding: 14px 20px;
  background-color: #F9FAFB;
  border-bottom: 1px solid #E5E7EB;
  color: #6B7280;
  font-size: 14px;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  text-align: left;
  white-space: nowrap;
}

.alm-table tbody td {
  padding: 20px;
  border-bottom: 1px solid #F1F5F9;
  color: #374151;
  vertical-align: middle;
}

.alm-table tbody tr:last-child td {
  border-bottom: 0;
}

.alm-table tbody tr:hover td {
  background-color: #F9FAFB;
}

.alm-cell-transmittal {
  color: #111827;
  font-weight: 600;
  white-space: nowrap;
}

.alm-person-name {
  color: #111827;
  font-weight: 500;
}

.alm-person-role {
  margin-top: 2px;
  color: #6B7280;
  font-size: 12.5px;
}

.alm-date {
  color: #374151;
  white-space: nowrap;
}

.alm-recipient {
  color: #374151;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.alm-doc-count {
  color: #111827;
  font-weight: 500;
  white-space: nowrap;
}

.alm-doc-name {
  margin-top: 2px;
  color: #6B7280;
  font-size: 12.5px;
  overflow-wrap: anywhere;
}

.alm-folder {
  min-width: 200px;
  max-width: 280px;
  color: #374151;
  line-height: 1.5;
  overflow-wrap: anywhere;
  word-break: break-word;
}

.alm-cell-action {
  width: 96px;
  text-align: right;
  white-space: nowrap;
}

.alm-cover-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  padding: 6px 14px;
  border: 1px solid #BFDBFE;
  border-radius: 8px;
  background-color: #ffffff;
  color: #1A73E8;
  font-family: inherit;
  font-size: 13px;
  font-weight: 500;
  line-height: 1.4;
  cursor: pointer;
  transition: background-color 150ms ease, border-color 150ms ease;
}

.alm-cover-btn:hover {
  background-color: #E8F0FE;
  border-color: #93C5FD;
}

.alm-cover-btn:focus-visible {
  outline: 2px solid #1A73E8;
  outline-offset: 2px;
}

/* ---------- Empty state ---------- */
.alm-empty {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 64px 24px;
  color: #6B7280;
  font-size: 15px;
  text-align: center;
}

.alm-visually-hidden {
  position: absolute !important;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

@media (min-width: 1024px) {
  .alm-overlay {
    align-items: center;
    padding: 40px 24px;
  }
}

@media (max-width: 640px) {
  .alm-header {
    padding: 18px;
    gap: 12px;
  }

  .alm-title {
    font-size: 20px;
  }

  .alm-header-icon {
    width: 40px;
    height: 40px;
    border-radius: 10px;
  }

  .alm-table thead th {
    padding: 12px 16px;
  }

  .alm-table tbody td {
    padding: 16px;
  }
}

@media (prefers-reduced-motion: reduce) {
  .alm-overlay,
  .alm-dialog,
  .alm-overlay--closing,
  .alm-dialog--closing {
    animation: none;
  }
}
`;

const MODAL_STYLES = MODAL_STYLES_PART_ONE + MODAL_STYLES_PART_TWO;

/** Adds the scoped stylesheet to <head> exactly once. */
const useModalStyles = (): void => {
    useEffect(() => {
        if (typeof document === "undefined") return;
        if (document.getElementById(STYLE_ELEMENT_ID)) return;

        const styleElement = document.createElement("style");
        styleElement.id = STYLE_ELEMENT_ID;
        styleElement.textContent = MODAL_STYLES;
        document.head.appendChild(styleElement);
    }, []);
};

/* ------------------------------------------------------------------ */
/* Modal                                                               */
/* ------------------------------------------------------------------ */
const AuditLogModal: React.FC<IAuditLogModalProps> = ({
    isOpen,
    entries,
    onRequestClose,
    onViewCover
}) => {
    useModalStyles();

    const [portalHost, setPortalHost] = useState<HTMLElement | null>(null);
    const [shouldRender, setShouldRender] = useState<boolean>(isOpen);
    const [isClosing, setIsClosing] = useState<boolean>(false);

    const dialogRef = useRef<HTMLDivElement>(null);
    const closeButtonRef = useRef<HTMLButtonElement>(null);
    const returnFocusRef = useRef<HTMLElement | null>(null);
    const exitTimerRef = useRef<number | undefined>(undefined);

    /* Portal target: <body>, so the popup is never clipped by the grid. */
    useEffect(() => {
        if (typeof document === "undefined") return;
        setPortalHost(document.body);
    }, []);

    /* Mount immediately, unmount only after the exit animation. */
    useEffect(() => {
        if (isOpen) {
            if (exitTimerRef.current !== undefined) {
                window.clearTimeout(exitTimerRef.current);
                exitTimerRef.current = undefined;
            }
            setIsClosing(false);
            setShouldRender(true);
            return;
        }

        if (!shouldRender) return;

        setIsClosing(true);
        exitTimerRef.current = window.setTimeout(() => {
            exitTimerRef.current = undefined;
            setIsClosing(false);
            setShouldRender(false);
        }, EXIT_ANIMATION_MS);
    }, [isOpen, shouldRender]);

    /* Clear a pending timer if the page unmounts while closing. */
    useEffect(() => () => {
        if (exitTimerRef.current !== undefined) {
            window.clearTimeout(exitTimerRef.current);
        }
    }, []);

    /* Lock page scroll while the popup is mounted (including its exit animation);
       keep the scrollbar gutter so nothing shifts. */
    useEffect(() => {
        if (!shouldRender || typeof document === "undefined") return;

        const body = document.body;
        const previousOverflow = body.style.overflow;
        const previousPaddingRight = body.style.paddingRight;
        const scrollbarWidth = window.innerWidth - document.documentElement.clientWidth;

        body.style.overflow = "hidden";
        if (scrollbarWidth > 0) {
            body.style.paddingRight = `${scrollbarWidth}px`;
        }

        return () => {
            body.style.overflow = previousOverflow;
            body.style.paddingRight = previousPaddingRight;
        };
    }, [shouldRender]);

    /* Focus in on open, back to the "Audit Log" toggle on close. */
    useEffect(() => {
        if (!isOpen || typeof document === "undefined") return;

        returnFocusRef.current = document.activeElement as HTMLElement | null;
        const frameId = window.requestAnimationFrame(() => {
            closeButtonRef.current?.focus();
        });

        return () => {
            window.cancelAnimationFrame(frameId);

            const toggle = document.querySelector("[data-testid=\"button-audit-log\"]") as HTMLElement | null;
            const previous = returnFocusRef.current;
            const target = previous && document.body.contains(previous) ? previous : toggle;
            if (target) target.focus();
        };
    }, [isOpen]);

    /* Escape closes the popup (capture phase, so Fluent panels can't eat it). */
    useEffect(() => {
        if (!isOpen || typeof document === "undefined") return;

        const handleKeyDown = (event: KeyboardEvent): void => {
            if (event.key !== "Escape") return;
            event.preventDefault();
            event.stopPropagation();
            onRequestClose();
        };

        document.addEventListener("keydown", handleKeyDown, true);
        return () => document.removeEventListener("keydown", handleKeyDown, true);
    }, [isOpen, onRequestClose]);

    /* Backdrop: a press that starts outside the dialog closes it. */
    const handleBackdropMouseDown = (event: React.MouseEvent<HTMLDivElement>): void => {
        if (event.target === event.currentTarget) {
            onRequestClose();
        }
    };

    /* Focus trap: keep Tab / Shift+Tab cycling inside the dialog. */
    const handleDialogKeyDown = (event: React.KeyboardEvent<HTMLDivElement>): void => {
        if (event.key !== "Tab") return;

        const dialog = dialogRef.current;
        if (!dialog) return;

        const nodes = dialog.querySelectorAll(FOCUSABLE_SELECTOR);
        const focusable: HTMLElement[] = [];
        for (let index = 0; index < nodes.length; index += 1) {
            const node = nodes[index] as HTMLElement;
            if (node.offsetWidth > 0 || node.offsetHeight > 0 || node === document.activeElement) {
                focusable.push(node);
            }
        }

        if (focusable.length === 0) {
            event.preventDefault();
            dialog.focus();
            return;
        }

        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        const active = document.activeElement;

        if (event.shiftKey && (active === first || active === dialog)) {
            event.preventDefault();
            last.focus();
            return;
        }

        if (!event.shiftKey && active === last) {
            event.preventDefault();
            first.focus();
        }
    };

    if (!portalHost || !shouldRender) return null;

    const lastEntry = entries.length > 0 ? entries[0] : null;

    return createPortal(
        <div
            className={`alm-overlay${isClosing ? " alm-overlay--closing" : ""}`}
            data-testid="overlay-audit-log"
            onMouseDown={handleBackdropMouseDown}
        >
            <div
                className={`alm-dialog${isClosing ? " alm-dialog--closing" : ""}`}
                role="dialog"
                aria-modal="true"
                aria-labelledby={TITLE_ID}
                tabIndex={-1}
                ref={dialogRef}
                onKeyDown={handleDialogKeyDown}
                onMouseDown={(event) => event.stopPropagation()}
            >
                <div className="alm-header">
                    <div className="alm-header-main">
                        <span className="alm-header-icon" aria-hidden="true">
                            <FileTextIcon />
                        </span>
                        <div className="alm-header-text">
                            <h2 id={TITLE_ID} className="alm-title">Transmittal Audit Log</h2>
                            {lastEntry ? (
                                <p className="alm-subtitle">
                                    Last transmittal <strong>{lastEntry.transmittalNo}</strong> sent by{" "}
                                    <strong>{lastEntry.sentBy}</strong> on{" "}
                                    <strong>{lastEntry.dateTime}</strong>.
                                </p>
                            ) : null}
                        </div>
                    </div>
                    <button
                        type="button"
                        ref={closeButtonRef}
                        className="alm-close"
                        aria-label="Close audit log"
                        title="Close audit log"
                        onClick={onRequestClose}
                    >
                        <XIcon />
                    </button>
                </div>

                <div className="alm-body">
                    {entries.length === 0 ? (
                        <div className="alm-empty" role="status">
                            No transmittals sent yet
                        </div>
                    ) : (
                        <table className="alm-table">
                            <caption className="alm-visually-hidden">
                                Transmittal audit log entries
                            </caption>
                            <thead>
                                <tr>
                                    <th scope="col">Transmittal</th>
                                    <th scope="col">Sent by</th>
                                    <th scope="col">Date &amp; time sent</th>
                                    <th scope="col">Recipient</th>
                                    <th scope="col">Documents</th>
                                    <th scope="col">Folder</th>
                                    <th scope="col">
                                        <span className="alm-visually-hidden">Actions</span>
                                    </th>
                                </tr>
                            </thead>
                            <tbody>
                                {entries.map((entry) => (
                                    <tr key={entry.id}>
                                        <td className="alm-cell-transmittal">{entry.transmittalNo}</td>
                                        <td>
                                            <div className="alm-person-name">{entry.sentBy}</div>
                                            {entry.role ? (
                                                <div className="alm-person-role">{entry.role}</div>
                                            ) : null}
                                        </td>
                                        <td className="alm-date">{entry.dateTime}</td>
                                        <td className="alm-recipient">{entry.recipient || "\u2014"}</td>
                                        <td>
                                            <div className="alm-doc-count">
                                                {formatDocumentCount(entry.documents.length)}
                                            </div>
                                            <div className="alm-doc-name">
                                                {entry.documents.length > 0
                                                    ? entry.documents.map((document) => document.name).join(", ")
                                                    : "\u2014"}
                                            </div>
                                        </td>
                                        <td>
                                            <div className="alm-folder">{entry.folderPath || "\u2014"}</div>
                                        </td>
                                        <td className="alm-cell-action">
                                            <button
                                                type="button"
                                                className="alm-cover-btn"
                                                onClick={() => onViewCover?.(entry)}
                                                title={`View the cover sheet for ${entry.transmittalNo}`}
                                            >
                                                Cover
                                            </button>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    )}
                </div>
            </div>
        </div>,
        portalHost
    );
};

export default React.memo(AuditLogModal);
