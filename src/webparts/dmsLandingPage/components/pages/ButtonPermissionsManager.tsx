import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    Checkbox,
    Toggle,
    PrimaryButton,
    DefaultButton,
    TextField,
    Spinner,
    SpinnerSize,
    MessageBar,
    MessageBarType,
    FontIcon,
} from '@fluentui/react';
import * as FluentIcons from '@fluentui/react-icons';
import { getAllButtonsAdmin, updateButtonItem, updateButtonsBatch } from './ButtonsService';
import PageLoader from '../../common/component/PageLoader';
import '../styles/ButtonPermissionsManager.css';

interface IButtonRow {
    ID: number;
    Title: string;
    InternalName: string;
    Active: boolean;
    Sequence: number;
    ButtonType: string;
    ButtonDisplayName: string;
    Icons: string;
    FullControl: boolean;
    Contribute: boolean;
    Edit: boolean;
    Read: boolean;
    // track local edits
    _dirty?: boolean;
    _saving?: boolean;
}

interface IButtonPermissionsManagerProps {
    context: WebPartContext;
}

const PERMISSION_COLS: { key: keyof IButtonRow; label: string; color: string; }[] = [
    { key: 'FullControl', label: 'Full Control', color: '#d13438' },
    { key: 'Contribute', label: 'Contribute', color: '#0078d4' },
    { key: 'Edit', label: 'Edit', color: '#107c10' },
    { key: 'Read', label: 'Read', color: '#8764b8' },
];

const BUTTON_TYPES = ['Document', 'Folder', 'Page', 'Other'];

/** Render a live Fluent UI icon from a string name */
function DynamicIcon({ name, size = 20 }: { name: string; size?: number; }): JSX.Element {
    const IconComp = FluentIcons[name as keyof typeof FluentIcons] as React.FC<{ style?: React.CSSProperties; }> | undefined;
    if (!IconComp) return <span style={{ fontSize: 12, color: '#a19f9d' }}>?</span>;
    return <IconComp style={{ width: size, height: size }} />;
}

/** Icon picker — a searchable dropdown listing common Fluent UI icon names */
const COMMON_ICONS = [
    'Eye20Regular', 'Delete20Regular', 'Share20Regular', 'Edit20Regular',
    'CheckmarkCircle20Regular', 'ArrowDownload20Regular', 'ArrowUpload20Regular',
    'DocumentPdf20Regular', 'DocumentText20Regular', 'Folder20Regular',
    'FolderOpen20Regular', 'Cloud20Regular', 'Settings20Regular',
    'People20Regular', 'Print20Regular', 'Save20Regular', 'Search20Regular',
    'Filter20Regular', 'Copy20Regular', 'LockClosed20Regular', 'LockOpen20Regular',
    'Add20Regular', 'Subtract20Regular', 'ChevronDown20Regular', 'Info20Regular',
    'Warning20Regular', 'ErrorCircle20Regular', 'Tag20Regular', 'Attach20Regular',
    'Link20Regular', 'History20Regular', 'Open20Regular', 'Navigation20Regular',
    'Home20Regular', 'Star20Regular', 'StarEmphasis20Regular', 'Pin20Regular',
    'Dismiss20Regular', 'DocumentBulletList20Regular', 'Table20Regular',
    'GridDots20Regular', 'List20Regular', 'CalendarMonth20Regular',
];

function IconPicker({ value, onChange }: { value: string; onChange: (v: string) => void; }): JSX.Element {
    const [open, setOpen] = useState(false);
    const [search, setSearch] = useState('');

    const filtered = COMMON_ICONS.filter(n => n.toLowerCase().includes(search.toLowerCase()));

    return (
        <div className="bpm-icon-picker" style={{ position: 'relative' }}>
            <div className="bpm-icon-preview" onClick={() => setOpen(o => !o)} title="Change icon">
                <DynamicIcon name={value} size={18} />
                <span className="bpm-icon-name">{value || '—'}</span>
                <FluentIcons.ChevronDown16Regular style={{ marginLeft: 4, color: '#605e5c' }} />
            </div>
            {open && (
                <div className="bpm-icon-dropdown">
                    <input
                        className="bpm-icon-search"
                        placeholder="Search icon..."
                        value={search}
                        onChange={e => setSearch(e.target.value)}
                        autoFocus
                    />
                    <div className="bpm-icon-grid">
                        {filtered.map(iconName => (
                            <div
                                key={iconName}
                                className={`bpm-icon-option ${value === iconName ? 'selected' : ''}`}
                                title={iconName}
                                onClick={() => { onChange(iconName); setOpen(false); setSearch(''); }}
                            >
                                <DynamicIcon name={iconName} size={20} />
                            </div>
                        ))}
                        {filtered.length === 0 && (
                            <div style={{ padding: '12px', color: '#a19f9d', fontSize: 12 }}>No icons found</div>
                        )}
                    </div>
                </div>
            )}
        </div>
    );
}

export default function ButtonPermissionsManager({ context }: IButtonPermissionsManagerProps): JSX.Element {
    const [rows, setRows] = useState<IButtonRow[]>([]);
    const [isLoading, setIsLoading] = useState(true);
    const [globalMsg, setGlobalMsg] = useState<{ text: string; type: MessageBarType; } | null>(null);
    const [searchText, setSearchText] = useState('');
    const [filterType, setFilterType] = useState<string>('All');
    const [savingAll, setSavingAll] = useState(false);

    /* ─── Load ─────────────────────────────────────────────── */
    const loadData = useCallback(async () => {
        setIsLoading(true);
        try {
            const res: any = await getAllButtonsAdmin(context);
            const items: IButtonRow[] = (res || []).map((item: any) => ({
                ID: item.ID,
                Title: item.Title || '',
                InternalName: item.InternalName || '',
                Active: !!item.Active,
                Sequence: item.Sequence ?? 0,
                ButtonType: item.ButtonType || '',
                ButtonDisplayName: item.ButtonDisplayName || '',
                Icons: item.Icons || '',
                FullControl: !!item.FullControl,
                Contribute: !!item.Contribute,
                Edit: !!item.EditPermission,
                Read: !!item.ReadPermission,
                _dirty: false,
                _saving: false,
            }));
            setRows(items);
        } catch (err) {
            console.error(err);
            setGlobalMsg({ text: 'Failed to load buttons. Check the DMS_Buttons list permissions.', type: MessageBarType.error });
        } finally {
            setTimeout(() => setIsLoading(false), 500); // Small buffer for smoother UX
        }
    }, [context]);

    useEffect(() => { void loadData(); }, [loadData]);

    /* ─── Helpers ───────────────────────────────────────────── */
    const updateRow = (id: number, partial: Partial<IButtonRow>) => {
        setRows(prev => prev.map(r => r.ID === id ? { ...r, ...partial, _dirty: true } : r));
    };

    const saveRow = async (row: IButtonRow) => {
        setRows(prev => prev.map(r => r.ID === row.ID ? { ...r, _saving: true } : r));
        try {
            await updateButtonItem(context, row.ID, {
                Title: row.Title,
                InternalName: row.InternalName,
                Active: row.Active,
                Sequence: row.Sequence,
                ButtonType: row.ButtonType,
                ButtonDisplayName: row.ButtonDisplayName,
                Icons: row.Icons,
                FullControl: row.FullControl,
                Contribute: row.Contribute,
                EditPermission: row.Edit,
                ReadPermission: row.Read,
            });
            setRows(prev => prev.map(r => r.ID === row.ID ? { ...r, _dirty: false, _saving: false } : r));
            setGlobalMsg({ text: `"${row.Title}" saved successfully.`, type: MessageBarType.success });
            setTimeout(() => setGlobalMsg(null), 3000);
        } catch (err) {
            console.error(err);
            setRows(prev => prev.map(r => r.ID === row.ID ? { ...r, _saving: false } : r));
            setGlobalMsg({ text: `Failed to save "${row.Title}". Please try again.`, type: MessageBarType.error });
        }
    };

    const saveAll = async () => {
        const dirty = rows.filter(r => r._dirty);
        if (dirty.length === 0) {
            setGlobalMsg({ text: 'No changes to save.', type: MessageBarType.info });
            setTimeout(() => setGlobalMsg(null), 2500);
            return;
        }
        setSavingAll(true);
        try {
            await updateButtonsBatch(context, dirty);
            setGlobalMsg({ text: `${dirty.length} item(s) saved successfully.`, type: MessageBarType.success });
            setRows(prev => prev.map(r => ({ ...r, _dirty: false, _saving: false })));
            setTimeout(() => setGlobalMsg(null), 3500);
        } catch (err) {
            console.error(err);
            setGlobalMsg({ text: 'Failed to save changes. Please try again.', type: MessageBarType.error });
        } finally {
            setSavingAll(false);
        }
    };

    /* ─── Filter ─────────────────────────────────────────────── */
    const filteredRows = rows.filter(r => {
        const matchesSearch =
            r.Title.toLowerCase().includes(searchText.toLowerCase()) ||
            r.InternalName.toLowerCase().includes(searchText.toLowerCase()) ||
            r.ButtonDisplayName.toLowerCase().includes(searchText.toLowerCase());
        const matchesType = filterType === 'All' || r.ButtonType === filterType;
        return matchesSearch && matchesType;
    });

    const dirtyCount = rows.filter(r => r._dirty).length;

    /* ─── Render ─────────────────────────────────────────────── */
    if (isLoading) {
        return (
            <PageLoader message="Loading Button Permissions..." minHeight="72vh" />
        );
    }

    return (
        <div className="bpm-page">
            {/* ── Page Header ── */}
            <div className="bpm-header">
                <div className="bpm-header-left">
                    <div className="bpm-header-icon-wrap">
                        <FluentIcons.ShieldLock24Regular className="bpm-header-icon" />
                    </div>
                    <div>
                        <h1 className="bpm-title">Button Permissions Manager</h1>
                        <p className="bpm-subtitle">
                            Configure visibility and access control for each action button across roles.
                        </p>
                    </div>
                </div>
                <div className="bpm-header-right">
                    {dirtyCount > 0 && (
                        <span className="bpm-dirty-badge">{dirtyCount} unsaved change{dirtyCount !== 1 ? 's' : ''}</span>
                    )}
                    <DefaultButton
                        text="Refresh"
                        iconProps={{ iconName: 'Refresh' }}
                        onClick={() => void loadData()}
                        styles={{ root: { borderRadius: 6 } }}
                    />
                    <PrimaryButton
                        text={savingAll ? 'Saving…' : 'Save All Changes'}
                        iconProps={{ iconName: 'Save' }}
                        disabled={savingAll || dirtyCount === 0}
                        onClick={() => void saveAll()}
                        styles={{ root: { borderRadius: 6 } }}
                    />
                </div>
            </div>

            {/* ── Status Bar ── */}
            {globalMsg && (
                <div className="bpm-msg-bar">
                    <MessageBar
                        messageBarType={globalMsg.type}
                        isMultiline={false}
                        onDismiss={() => setGlobalMsg(null)}
                        dismissButtonAriaLabel="Close"
                    >
                        {globalMsg.text}
                    </MessageBar>
                </div>
            )}

            {/* ── Toolbar ── */}
            <div className="bpm-toolbar">
                <div className="bpm-toolbar-left">
                    <TextField
                        placeholder="Search by title, internal name, or display name..."
                        value={searchText}
                        onChange={(_, v) => setSearchText(v || '')}
                        styles={{
                            root: { width: 340 },
                            field: { fontSize: 13 },
                            fieldGroup: { borderRadius: 6, height: 36 }
                        }}
                        iconProps={{ iconName: 'Search' }}
                    />
                    <div className="bpm-type-filters">
                        {['All', ...BUTTON_TYPES].map(t => (
                            <button
                                key={t}
                                className={`bpm-type-chip ${filterType === t ? 'active' : ''}`}
                                onClick={() => setFilterType(t)}
                            >
                                {t}
                            </button>
                        ))}
                    </div>
                </div>
                <span className="bpm-count-label">{filteredRows.length} button{filteredRows.length !== 1 ? 's' : ''}</span>
            </div>

            {/* ── Main Table ── */}
            <div className="bpm-table-wrapper">
                <table className="bpm-table">
                    <thead>
                        <tr>
                            <th className="bpm-th bpm-th-seq">Seq.</th>
                            <th className="bpm-th">Title</th>
                            <th className="bpm-th">Display Name</th>
                            <th className="bpm-th bpm-th-icon">Icon</th>
                            <th className="bpm-th bpm-th-toggle">Active</th>
                            {/* Permission matrix columns */}
                            {PERMISSION_COLS.map(col => (
                                <th key={col.key as string} className="bpm-th bpm-th-perm" style={{ '--perm-color': col.color } as React.CSSProperties}>
                                    <span className="bpm-perm-label" style={{ color: col.color }}>{col.label}</span>
                                </th>
                            ))}
                            <th className="bpm-th bpm-th-action">Action</th>
                        </tr>
                    </thead>
                    <tbody>
                        {filteredRows.length === 0 && (
                            <tr>
                                <td colSpan={10} className="bpm-empty-row">
                                    <FluentIcons.DocumentSearch24Regular style={{ color: '#c8c6c4', marginBottom: 6 }} />
                                    <span>No buttons match your search.</span>
                                </td>
                            </tr>
                        )}
                        {filteredRows.map((row, idx) => (
                            <tr
                                key={row.ID}
                                className={`bpm-tr ${row._dirty ? 'bpm-tr-dirty' : ''} ${row._saving ? 'bpm-tr-saving' : ''}`}
                            >
                                {/* Sequence number display */}
                                <td className="bpm-td bpm-td-seq">
                                    <span className="bpm-seq-badge">{idx + 1}</span>
                                </td>

                                {/* Title */}
                                <td className="bpm-td">
                                    <span className="bpm-cell-text bpm-title-text">{row.Title}</span>
                                </td>

                                {/* Display Name */}
                                <td className="bpm-td">
                                    <input
                                        type="text"
                                        className="bpm-text-input"
                                        value={row.ButtonDisplayName}
                                        onChange={e => updateRow(row.ID, { ButtonDisplayName: e.target.value })}
                                        placeholder="Display name..."
                                    />
                                </td>

                                {/* Icon Picker */}
                                <td className="bpm-td bpm-td-icon">
                                    <IconPicker
                                        value={row.Icons}
                                        onChange={v => updateRow(row.ID, { Icons: v })}
                                    />
                                </td>

                                {/* Active Toggle */}
                                <td className="bpm-td bpm-td-toggle">
                                    <Toggle
                                        checked={row.Active}
                                        onChange={(_, checked) => updateRow(row.ID, { Active: !!checked })}
                                        styles={{
                                            root: { marginBottom: 0 },
                                            pill: { background: row.Active ? '#107c10' : '#c8c6c4' }
                                        }}
                                    />
                                </td>

                                {/* Permission checkboxes */}
                                {PERMISSION_COLS.map(col => (
                                    <td key={col.key as string} className="bpm-td bpm-td-perm">
                                        <div className="bpm-perm-cell">
                                            <Checkbox
                                                checked={!!row[col.key]}
                                                onChange={(_, checked) => updateRow(row.ID, { [col.key]: !!checked } as Partial<IButtonRow>)}
                                                styles={{
                                                    checkbox: {
                                                        borderColor: row[col.key] ? col.color : '#c8c6c4',
                                                        background: row[col.key] ? col.color : 'transparent',
                                                    },
                                                    checkmark: { color: '#fff' }
                                                }}
                                            />
                                        </div>
                                    </td>
                                ))}

                                {/* Row action */}
                                <td className="bpm-td bpm-td-action">
                                    {row._saving ? (
                                        <Spinner size={SpinnerSize.small} />
                                    ) : (
                                        <button
                                            className={`bpm-save-row-btn ${row._dirty ? 'has-changes' : ''}`}
                                            onClick={() => void saveRow(row)}
                                            disabled={!row._dirty}
                                            title={row._dirty ? 'Save this row' : 'No changes'}
                                        >
                                            <FluentIcons.Save20Regular />
                                            <span>Save</span>
                                        </button>
                                    )}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* ── Permission Legend ── */}
            <div className="bpm-legend">
                <span className="bpm-legend-title">Permission Levels:</span>
                {PERMISSION_COLS.map(col => (
                    <div key={col.key as string} className="bpm-legend-item">
                        <span className="bpm-legend-dot" style={{ background: col.color }} />
                        <span className="bpm-legend-label">{col.label}</span>
                    </div>
                ))}
                <span className="bpm-legend-sep">|</span>
                <span className="bpm-legend-hint">
                    <span className="bpm-dirty-indicator" /> Unsaved row
                </span>
            </div>
        </div>
    );
}
