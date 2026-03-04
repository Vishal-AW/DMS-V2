import * as React from "react";
import {
    AllCommunityModule,
    ModuleRegistry,
    themeQuartz,
} from "ag-grid-community";
import { AgGridReact } from "ag-grid-react";
// import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-alpine.css";
ModuleRegistry.registerModules([AllCommunityModule]);
import { useEffect, useRef } from "react";
//import { useTheme } from "office-ui-fabric-react";
import type { ColDef } from "ag-grid-community";

export interface IReusableDataTableComponentProps {
    rowData: any[];
    columnDefs: ColDef[];
    loading?: boolean;
    onGridReady?: (params: any) => void;
    searchText?: string;
}

const ReusableDataTable: React.FC<IReusableDataTableComponentProps> = ({
    rowData,
    columnDefs,
    loading,
    onGridReady,
    searchText,
}) => {
    const gridRef = useRef<AgGridReact<any>>(null);
    //const theme = useTheme();
    const myTheme = themeQuartz.withParams({
        headerTextColor: "#2563EB",
        headerBackgroundColor: "#EFF6FF",
        rowHoverColor: "#DBEAFE",
    });

    const defaultColDef = React.useMemo(
        () => ({
            editable: false,
            filter: true,
            flex: 1,
            minWidth: 100,
        }),
        [],
    );

    useEffect(() => {
        if (gridRef.current) {
            (gridRef?.current?.api as any)?.setQuickFilter?.(searchText);
        }
    }, [searchText]);
    return (
        <div className="ag-theme-balham" style={{ width: "100%", height: "600px" }}>
            <AgGridReact
                ref={gridRef}
                rowData={rowData}
                theme={myTheme}
                columnDefs={columnDefs}
                defaultColDef={defaultColDef}
                paginationAutoPageSize={true}
                pagination={true}
                rowSelection="multiple"
                onGridReady={onGridReady}
                loading={loading}
                quickFilterText={searchText}
                multiSortKey="ctrl"
            />
        </div>
    );
};

export default React.memo(ReusableDataTable);
