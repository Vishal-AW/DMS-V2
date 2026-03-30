/* eslint-disable */
import * as React from 'react';
import { AllCommunityModule, ModuleRegistry, themeQuartz, RowDragModule } from "ag-grid-community";
import { AgGridReact } from "ag-grid-react";
import { useEffect, useMemo, useRef } from "react";
import type { ColDef } from "ag-grid-community";
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
}

const ReusableDataTable: React.FC<IReusableDataTableComponentProps> = ({
    rowData, columnDefs, loading, onGridReady, searchText, onRowDragEnd, pagination = true
}) => {
    const gridRef = useRef<AgGridReact<any>>(null);
    const [filterData, setFilterData] = React.useState<any[]>(rowData);

    const agTheme = useMemo(
        () =>
            themeQuartz.withParams({
                headerTextColor: "#2563EB",
                headerBackgroundColor: "#EFF6FF",
                rowHoverColor: "#DBEAFE",
                oddRowBackgroundColor: tokens.colorNeutralBackground1,
                // rowBorderColor: tokens.colorNeutralStroke2,
            }),
        []
    );

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
        <div style={{ width: "100%", height: "600px" }} className="ag-theme-alpine">
            <AgGridReact
                ref={gridRef}
                rowData={filterData}
                columnDefs={columns}
                theme={agTheme}
                defaultColDef={defaultColDef}
                pagination={pagination}
                paginationPageSize={10}
                paginationPageSizeSelector={[10, 20, 50, 100]}
                rowSelection="multiple"
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
