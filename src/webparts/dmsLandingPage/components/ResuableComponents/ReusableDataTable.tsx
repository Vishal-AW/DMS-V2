/* eslint-disable */
import * as React from 'react';
import { AllCommunityModule, ModuleRegistry, themeQuartz, RowDragModule } from "ag-grid-community";
import { AgGridReact } from "ag-grid-react";
import { useCallback, useEffect, useMemo, useRef } from "react";
import type { ColDef, MultiRowSelectionOptions, SelectionColumnDef } from "ag-grid-community";
import { tokens } from "@fluentui/react-components";

ModuleRegistry.registerModules([AllCommunityModule, RowDragModule]);

export interface IReusableDataTableComponentProps {
    rowData: any[];
    columnDefs: ColDef[];
    loading?: boolean;
    onGridReady?: (params: any) => void;
    searchText?: string;
    onRowDragEnd?: (event: any) => void;
    pagination?: boolean;
    /* ------------------------------------------------------------------ */
    /* Optional row selection (design feature - only used when enabled).   */
    /* Leaving these props out keeps the previous grid behaviour exactly.  */
    /* ------------------------------------------------------------------ */
    /** Renders a checkbox selection column + select-all header checkbox. */
    selectionEnabled?: boolean;
    /** Stable id per row so the selection survives sorting / filtering. */
    getRowKey?: (data: any) => string;
    /** Receives the currently selected row objects. */
    onSelectionChange?: (selectedRows: any[]) => void;
    /** When this value changes the selection is cleared (e.g. folder change). */
    selectionResetKey?: string | number;
    /** Screen-reader label for a row checkbox, e.g. "Select Invoice.docx". */
    getRowAriaLabel?: (data: any) => string;
}

const ReusableDataTable: React.FC<IReusableDataTableComponentProps> = ({
    rowData, columnDefs, loading, onGridReady, searchText, onRowDragEnd, pagination = true,
    selectionEnabled = false, getRowKey, onSelectionChange, selectionResetKey, getRowAriaLabel
}) => {
    const gridRef = useRef<AgGridReact<any>>(null);
    const containerRef = useRef<HTMLDivElement>(null);
    const [filterData, setFilterData] = React.useState<any[]>(rowData);

    const agTheme = useMemo(
        () =>
            themeQuartz.withParams({
                headerTextColor: "#2563EB",
                headerBackgroundColor: "#EFF6FF",
                rowHoverColor: selectionEnabled ? "#f8fafc" : "#DBEAFE",
                oddRowBackgroundColor: tokens.colorNeutralBackground1,
                // Extra design tokens - only for the checkbox (document) grid so
                // every other table keeps its current appearance.
                ...(selectionEnabled
                    ? {
                        rowHeight: 76,
                        headerHeight: 48,
                        headerTextColor: "#475569",
                        selectedRowBackgroundColor: "#e8f0fe",
                        checkboxBorderRadius: 4,
                        checkboxBorderWidth: 1,
                        checkboxUncheckedBackgroundColor: "#ffffff",
                        checkboxUncheckedBorderColor: "#cbd5e1",
                        checkboxCheckedBackgroundColor: "#1a6fe8",
                        checkboxCheckedBorderColor: "#1a6fe8",
                        checkboxCheckedShapeColor: "#ffffff",
                        checkboxIndeterminateBackgroundColor: "#1a6fe8",
                        checkboxIndeterminateBorderColor: "#1a6fe8",
                        checkboxIndeterminateShapeColor: "#ffffff",
                    }
                    : {}),
            }),
        [selectionEnabled]
    );

    /* ------------------------------------------------------------------ */
    /* Row selection (opt-in)                                              */
    /* ------------------------------------------------------------------ */
    const rowSelection = useMemo<MultiRowSelectionOptions>(
        () => ({
            mode: "multiRow",
            checkboxes: true,
            headerCheckbox: true,
            // Only the checkbox selects a row, so the existing cell click
            // handlers (open document / action menu) keep working unchanged.
            enableClickSelection: false,
            // The header checkbox selects / clears the rows on the current page.
            selectAll: "currentPage",
        }),
        []
    );

    const selectionColumn = useMemo<SelectionColumnDef>(
        () => ({
            width: 52,
            minWidth: 52,
            maxWidth: 52,
            pinned: "left",
            lockPosition: "left",
            resizable: false,
            sortable: false,
            suppressHeaderMenuButton: true,
        }),
        []
    );

    const handleSelectionChanged = useCallback((event: any) => {
        if (!onSelectionChange) return;
        onSelectionChange(event?.api?.getSelectedRows?.() ?? []);
    }, [onSelectionChange]);

    const rowIdGetter = useCallback((params: any) => String(getRowKey ? getRowKey(params.data) : params.data?.id ?? ""), [getRowKey]);

    // Clear the selection whenever the consumer changes the reset key
    // (the workspace passes the current folder path).
    useEffect(() => {
        if (!selectionEnabled) return;
        gridRef.current?.api?.deselectAll();
    }, [selectionEnabled, selectionResetKey]);

    /** Adds a meaningful accessible name to the row checkboxes. */
    const applyRowCheckboxLabels = useCallback(() => {
        if (!selectionEnabled || !getRowAriaLabel) return;
        try {
            const api = gridRef.current?.api;
            const root = containerRef.current;
            if (!api || !root) return;
            const rowElements = root.querySelectorAll<HTMLElement>(".ag-row");
            for (let i = 0; i < rowElements.length; i++) {
                const rowElement = rowElements[i];
                const rowIndex = Number(rowElement.getAttribute("row-index"));
                if (isNaN(rowIndex)) continue;
                const node = api.getDisplayedRowAtIndex(rowIndex);
                if (!node || !node.data) continue;
                const label = getRowAriaLabel(node.data);
                if (!label) continue;
                const checkbox = rowElement.querySelector<HTMLInputElement>("input.ag-checkbox-input");
                if (checkbox && checkbox.getAttribute("aria-label") !== label) {
                    checkbox.setAttribute("aria-label", label);
                }
            }
        } catch {
            /* Accessibility polish only - never block rendering on this. */
        }
    }, [selectionEnabled, getRowAriaLabel]);

    const scheduleRowCheckboxLabels = useCallback(() => {
        if (!selectionEnabled) return;
        window.setTimeout(applyRowCheckboxLabels, 0);
    }, [selectionEnabled, applyRowCheckboxLabels]);

    useEffect(() => {
        scheduleRowCheckboxLabels();
    }, [scheduleRowCheckboxLabels, filterData]);

    const defaultColDef = useMemo(
        () => ({
            editable: false,
            filter: true,
            flex: 1,
            minWidth: 120,
        }),
        []
    );

    useEffect(() => {
        if (!searchText) setFilterData(rowData);
        else if (searchText.trim()) {
            const query = searchText.toLowerCase();
            const result = rowData.filter((item) =>
                Object.values(item).some(
                    (value) =>
                        value !== null &&
                        value !== undefined &&
                        value
                            .toString()
                            .toLowerCase()
                            .includes(query)
                )
            );
            setFilterData(result);
        }
    }, [searchText, rowData]);

    // const columns = columnDefs.map(col => ({
    //     ...col,
    //     headerName: col.headerName?.toUpperCase() || col.field?.toUpperCase()
    // }));

    const toPrettyTitle = (text: string) =>
        text
            .replace(/_/g, " ")
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .toLowerCase()
            .replace(/\b\w/g, c => c.toUpperCase());

    const columns = columnDefs.map((col, index) => ({
        ...col,
        headerName: toPrettyTitle(col.headerName || col.field || ""),
        maxWidth: index === 0 ? 80 : undefined
    }));


    return (
        <div
            ref={containerRef}
            style={{ width: "100%", ...(selectionEnabled ? {} : { height: "600px" }) }}
            className={selectionEnabled ? "ag-theme-alpine reusable-data-table-selection" : "ag-theme-alpine"}
        >
            <AgGridReact
                ref={gridRef}
                rowData={filterData}
                columnDefs={columns}
                theme={agTheme}
                defaultColDef={defaultColDef}
                pagination={pagination}
                paginationPageSize={10}
                paginationPageSizeSelector={[10, 20, 50, 100]}
                rowSelection={selectionEnabled ? rowSelection : "multiple"}
                selectionColumnDef={selectionEnabled ? selectionColumn : undefined}
                getRowId={selectionEnabled && getRowKey ? rowIdGetter : undefined}
                onSelectionChanged={selectionEnabled ? handleSelectionChanged : undefined}
                onModelUpdated={selectionEnabled ? scheduleRowCheckboxLabels : undefined}
                localeText={selectionEnabled ? { ariaRowSelectAll: "Select or clear all documents on this page" } : undefined}
                onGridReady={onGridReady}
                loading={loading}
                multiSortKey="ctrl"
                rowDragManaged
                onRowDragEnd={onRowDragEnd}
                animateRows
            />
        </div>
    );
};

export default React.memo(ReusableDataTable);
