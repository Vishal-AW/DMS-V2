import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useEffect, useState } from 'react';
import ReusableDataTable from "../ResuableComponents/ReusableDataTable";
import { ILabel } from "../../../../Intrface/ILabel";
import { SearchBox } from "@fluentui/react";
import { getTileAllData } from "../../../../Services/MasTileService";
import { Badge, Button } from "@fluentui/react-components";
import { Add20Regular, Edit24Regular } from "@fluentui/react-icons";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { format } from "date-fns";
import TileForm from "./TileForm";
import PageLoader from "../../common/component/PageLoader";
interface ITileSettingProps {
    context: WebPartContext;
}

const TileSetting: React.FunctionComponent<ITileSettingProps> = ({ context }) => {
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const [rowData, setRowData] = useState<any[]>([]);
    const [searchQuery, setSearchQuery] = useState<string>("");
    const SiteURL = context.pageContext.web.absoluteUrl;
    const [isOpenEditor, setIsOpenEditor] = useState<boolean>(false);
    const [itemId, setItemId] = useState<number>(0);
    const [isLoading, setIsLoading] = useState<boolean>(true);

    useEffect(() => {
        setIsLoading(true);
        fetchTileData();
    }, [isOpenEditor]);

    const fetchTileData = async () => {
        let FetchallTileData: any = await getTileAllData(SiteURL, context.spHttpClient);
        let TilesData = FetchallTileData.value;
        setRowData(TilesData);
        setIsLoading(false);
    };

    const columns = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo,
                filter: false,
                resizable: false,
                maxWidth: 80,
                field: "Order0",
                rowDrag: true
            },
            {
                headerName: DisplayLabel.Tiles,
                filter: true,
                sortable: true,
                field: "TileName",
            },
            {
                headerName: DisplayLabel.AllowApprover,
                filter: true,
                sortable: true,
                field: "AllowApprover",
                cellRenderer: (item: any) => (item.data.AllowApprover ? "Yes" : "No")
            },
            {
                headerName: DisplayLabel.LastModified,
                filter: true,
                sortable: true,
                field: "Modified",
                cellRenderer: (item: any) => {
                    const formattedDate = format(item.data.Modified, "dd/MM/yyyy");
                    const formattedTime = new Date(item.data.Modified).toLocaleTimeString("en-US", {
                        hour: "2-digit",
                        minute: "2-digit",
                        hour12: true
                    });
                    return `${item.data.Editor?.Title} ${formattedDate} at ${formattedTime}`;
                }
            },
            {
                headerName: DisplayLabel.Status,
                filter: true,
                sortable: true,
                field: "Active",
                cellRenderer: (item: any) =>
                    <div style={{
                        display: "flex",
                        alignItems: "center",
                        gap: "8px"
                    }}>
                        <Badge appearance="filled" color={item.data.Active === true ? "success" : "informative"} />
                        {item.data.Active === true ? "Yes" : "No"}
                    </div>
            },
            {
                headerName: DisplayLabel.Action,
                filter: false,
                sortable: false,
                cellRenderer: (item: any) => <div className="flex gap-xs">
                    <Button
                        appearance="subtle"
                        icon={<Edit24Regular />}
                        size="small"
                        onClick={() => openEditEditor(item.data)}
                        title="Edit"
                        style={{ color: "#009ef7" }}
                        data-testid={`button-edit-page-${item.data.ID}`}
                    />
                </div>
            }
        ];
    }, []);

    const openEditEditor = async (item: any) => {
        setIsOpenEditor(true);
        setItemId(item?.ID);
    };

    const onRowDragEnd = async (event: any) => {
        const movedItem = event.node.data;
        const fromIndex = movedItem.Order0;
        const toIndex = event.overIndex + 1;
        if (fromIndex === toIndex) return;
        const direction = fromIndex > toIndex ? "forward" : "backward";
        const updatedList = await UpdateSequenceNumber(fromIndex, toIndex, direction);
        setRowData(updatedList);
    };

    const UpdateSequenceNumber = async (fromIndex: number, toIndex: number, direction: string) => {
        let newList = [...rowData].map(item => ({ ...item }));
        const movedItem = newList.find(i => i.Order0 === fromIndex);
        if (!movedItem) return newList;

        newList = newList.filter(i => i.Order0 !== fromIndex);
        if (direction === "forward") {
            newList.forEach(item => {
                if (item.Order0 >= toIndex && item.Order0 < fromIndex)
                    item.Order0 += 1;
            });
        }
        else {
            newList.forEach(item => {
                if (item.Order0 <= toIndex && item.Order0 > fromIndex)
                    item.Order0 -= 1;
            });
        }

        movedItem.Order0 = toIndex;
        newList.push(movedItem);
        newList.sort((a, b) => a.Order0 - b.Order0);
        await updateSequenceInSP(newList);
        return newList;
    };

    const updateSequenceInSP = async (items: any[]) => {
        const sp = spfi().using(SPFx(context));
        const [batchedSP, execute] = sp.web.batched();

        const list = batchedSP.lists.getByTitle("DMS_Mas_Tile");

        items.forEach(item => {
            list.items.getById(item.Id).update({
                Order0: item.Order0,
            });
        });
        await execute();
    };

    if (isOpenEditor) {
        return <TileForm context={context} tileID={itemId} setIsOpenEditor={setIsOpenEditor} allTiles={rowData} />;
    }

    if (isLoading) {
        return <PageLoader message="Loading workspace tiles..." minHeight="72vh" />;
    }

    return (
        <div className="tile-settings-page" data-testid="page-tile-settings">
            <div className="tile-settings-body">
                <div className="tile-settings-toolbar">
                    <h2 className="tile-settings-subtitle">Workspace Tiles</h2>
                    <div className="tile-settings-search">
                        <SearchBox
                            placeholder="Search..."
                            value={searchQuery}
                            onChange={(_, value) => setSearchQuery(value || '')}
                            onClear={() => setSearchQuery('')}
                            className="dashboard-search-box"
                            data-testid="input-search-tiles"
                        />
                    </div>
                    <>
                        <Button
                            appearance="primary"
                            icon={<Add20Regular />}
                            size="medium"
                            onClick={() => { setIsOpenEditor(true); setItemId(0); }}
                            title="Edit"
                        >{DisplayLabel.Add}</Button>
                    </>
                </div>
                <div className="tile-settings-table-wrap" data-testid="table-tiles">
                    <ReusableDataTable rowData={rowData} columnDefs={columns} onRowDragEnd={onRowDragEnd} searchText={searchQuery} pagination={false} />
                </div>
            </div>
        </div>
    );
};

export default React.memo(TileSetting);;
