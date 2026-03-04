import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { DefaultButton, PrimaryButton, IconButton, IContextualMenuProps, IContextualMenuItem } from '@fluentui/react';
import { ArrowUpload20Regular, FolderAdd20Regular, Add20Regular, Home20Regular, ChevronRight12Regular, Filter20Regular } from '@fluentui/react-icons';
import Sidebar from "../../common/component/Sidebar";
import { FolderNode } from "../../common/component/FolderTree";
import { buildBreadcrumbPath, buildFolderHierarchy, buildLibraryRootPath, getAllDocuments } from "../../common/commonfunction";
import ReusableDataTable from "../ResuableComponents/ReusableDataTable";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ILabel } from "../../../../Intrface/ILabel";

interface IWorkspaceProps {
    context: WebPartContext;
}

const Workspace: React.FunctionComponent<IWorkspaceProps> = ({ context }) => {
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const { workspaceId } = useParams<{ workspaceId: string; }>();
    const navigate = useNavigate();
    const [selectedFolder, setSelectedFolder] = useState<FolderNode | null>(null);
    // const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [folders, setFolders] = useState<any>([]);
    const [tileData, setTileData] = useState<any | null>(null);
    const [files, setFiles] = useState<any[]>([]);

    useEffect(() => {
        fetchTileData();
    }, []);
    const fetchTileData = async () => {
        const sp = spfi().using(SPFx(context));
        const data = await sp.web.lists.getByTitle("DMS_Mas_Tile").items.getById(Number(workspaceId))();
        setTileData(data);
    };

    useEffect(() => {
        tileData && fetchFolder();
    }, [tileData]);

    const fetchFolder = async () => {
        const sp = spfi().using(SPFx(context));

        const allFolders: any[] = [];

        const items = await sp.web.lists
            .getByTitle(tileData?.LibraryName)
            .items
            .select("Id", "Title", "FileRef", "FileDirRef", "FSObjType")
            .filter("FSObjType eq 1")
            .top(5000);

        for await (const batch of items) {
            allFolders.push(...batch);
        }

        const rootPath = buildLibraryRootPath(context, tileData?.LibraryName);
        const folder = buildFolderHierarchy(allFolders, rootPath);
        setFolders([{
            id: 0,
            name: tileData?.LibraryName,
            path: "",
            children: [...folder]
        }]);
    };

    const handleFolderSelect = (folder: FolderNode) => {
        setSelectedFolder(folder);
    };

    const handleFolderAction = (action: string, folder: FolderNode) => {
        console.log('Folder action:', action, folder.name);
        // todo: implement folder actions via SharePoint API
    };

    const handleNewFolder = () => {
        // setIsPanelOpen(true);
    };

    const handleCreateFolder = (data: Record<string, any>) => {
        console.log('Creating folder with data:', data);
        // todo: implement SharePoint folder creation
    };

    const handleDocumentAction = (action: string, doc: any) => {
        console.log('Document action:', action, doc.name);
        // todo: implement document actions via SharePoint API
    };

    const handleUpload = () => {
        console.log('Upload triggered');
        // todo: implement file upload
    };

    const folderPath = useMemo<FolderNode[]>(() => {
        if (!selectedFolder) return [];
        return buildBreadcrumbPath(selectedFolder, folders);
    }, [selectedFolder]);

    useEffect(() => {
        getDocument();
    }, [selectedFolder]);

    const getDocument = async () => {
        if (!selectedFolder) return [];
        if (selectedFolder.isLastLevel) {
            const files = await getAllDocuments(context, selectedFolder?.path);
            setFiles(files);
        };
    };
    const onAction = (action: any, doc: any) => {

    };
    const getActionMenuProps = (doc: any): IContextualMenuProps => {
        const items: IContextualMenuItem[] = [
            {
                key: 'history',
                text: 'History',
                iconProps: { iconName: 'History' },
                onClick: () => onAction?.('history', doc),
            },
            {
                key: 'view',
                text: 'View',
                iconProps: { iconName: 'View' },
                onClick: () => onAction?.('view', doc),
            },
            {
                key: 'preview',
                text: 'Preview',
                iconProps: { iconName: 'EntryView' },
                onClick: () => onAction?.('preview', doc),
            },
            {
                key: 'download',
                text: 'Download',
                iconProps: { iconName: 'Download' },
                onClick: () => onAction?.('download', doc),
            },
            {
                key: 'rename',
                text: 'Rename',
                iconProps: { iconName: 'Rename' },
                onClick: () => onAction?.('rename', doc),
            },
            {
                key: 'versions',
                text: 'Versions',
                iconProps: { iconName: 'History' },
                onClick: () => onAction?.('versions', doc),
            },
            {
                key: 'checkout',
                text: 'Check out',
                iconProps: { iconName: 'PageCheckedOut' },
                onClick: () => onAction?.('checkout', doc),
            },
            {
                key: 'share',
                text: 'Share',
                iconProps: { iconName: 'Share' },
                onClick: () => onAction?.('share', doc),
            },
            {
                key: 'permissions',
                text: 'Advance Permission',
                iconProps: { iconName: 'Permissions' },
                onClick: () => onAction?.('permissions', doc),
            },
        ];

        return { items };
    };

    const columns = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo,
                filter: false,
                resizable: false,
                valueGetter: (params: any) => params.node.rowIndex + 1
            },
            {
                headerName: DisplayLabel.FileName,
                filter: true,
                sortable: true,
                field: "Name"
            },
            {
                headerName: DisplayLabel.ReferenceNo,
                filter: true,
                sortable: true,
                field: "ReferenceNo",
            },
            {
                headerName: DisplayLabel.Versions,
                filter: true,
                sortable: true,
                field: "Level",
                cellRenderer: (item: any) =>
                    <span className="table-cell-text table-cell-version" data-testid={`text-version-${item.id}`}>
                        v{item?.data?.Level}
                    </span>
            },
            {
                headerName: DisplayLabel.Status,
                filter: true,
                sortable: true,
                field: "DisplayStatus",
            },
            {
                headerName: DisplayLabel.Action,
                filter: true,
                sortable: true,
                cellRenderer: (item: any) => <IconButton
                    menuProps={getActionMenuProps(item)}
                    iconProps={{ iconName: 'More' }}
                    title="Actions"
                    ariaLabel="Document actions"
                    className="table-action-btn"
                    data-testid={`button-actions-${item.id}`}
                />
            }
        ];
    }, []);

    return (
        <div className="workspace-page" data-testid="page-workspace-explorer">
            <div className="workspace-topbar">
                <div className="workspace-topbar-breadcrumb" data-testid="nav-top-breadcrumb">
                    <span
                        className="workspace-topbar-link"
                        onClick={() => navigate('/')}
                        data-testid="link-dashboard"
                    >
                        <Home20Regular className="workspace-topbar-home-icon" />
                        <span>Dashboard</span>
                    </span>
                    <ChevronRight12Regular className="workspace-topbar-separator-icon" />
                    <span className="workspace-topbar-current" data-testid="text-workspace-name">
                        {tileData?.TileName}
                    </span>
                </div>
                <div className="workspace-topbar-actions">
                    <DefaultButton
                        className="workspace-new-request-btn"
                        onClick={() => console.log('New Request clicked')}
                        data-testid="button-new-request"
                    >
                        <Add20Regular className="workspace-btn-icon" />
                        <span>New Request</span>
                    </DefaultButton>
                </div>
            </div>

            <div className="workspace-body">
                <Sidebar
                    folders={folders}
                    selectedFolderId={selectedFolder ? selectedFolder.id : "0"}
                    onFolderSelect={handleFolderSelect}
                    onFolderAction={handleFolderAction}
                    recycleBinCount={0}
                    approvalCount={0}
                    onRecycleBinClick={() => console.log('Recycle Bin clicked')}
                    onApprovalClick={() => navigate('/approvals', { state: { from: `/workspace/${workspaceId}` } })}
                    onAdvancedSearchClick={() => navigate('/search', { state: { from: `/workspace/${workspaceId}` } })}
                />

                <div className="workspace-content">
                    {selectedFolder && folderPath.length > 0 && (
                        <div className="workspace-content-header">
                            <div className="workspace-folder-breadcrumb" data-testid="nav-folder-breadcrumb">
                                {folderPath.map((node, i) => (
                                    <span key={node.id} className="workspace-folder-breadcrumb-segment">
                                        {i > 0 && <ChevronRight12Regular className="workspace-folder-breadcrumb-chevron" />}
                                        <span
                                            className={`workspace-folder-breadcrumb-item ${i === folderPath.length - 1 ? 'workspace-folder-breadcrumb-current' : ''}`}
                                            onClick={() => {
                                                if (i < folderPath.length - 1) handleFolderSelect(node);
                                            }}
                                            data-testid={`breadcrumb-folder-${node.id}`}
                                        >
                                            {node.name}
                                        </span>
                                    </span>
                                ))}
                            </div>
                            <div className="workspace-content-actions">
                                <IconButton
                                    className="workspace-action-icon-btn"
                                    title="Filter"
                                    ariaLabel="Filter"
                                    data-testid="button-filter"
                                >
                                    <Filter20Regular />
                                </IconButton>
                                <DefaultButton
                                    onClick={handleUpload}
                                    className="workspace-upload-btn"
                                    data-testid="button-upload"
                                >
                                    <ArrowUpload20Regular className="workspace-btn-icon" />
                                    <span>Upload</span>
                                </DefaultButton>
                                <PrimaryButton
                                    onClick={handleNewFolder}
                                    className="workspace-new-folder-btn"
                                    data-testid="button-new-folder"
                                >
                                    <FolderAdd20Regular className="workspace-btn-icon" />
                                    <span>New Folder</span>
                                </PrimaryButton>
                            </div>
                        </div>
                    )}

                    {selectedFolder ? (
                        <ReusableDataTable rowData={files} columnDefs={[]} />
                        // <></>
                    ) : (
                        <div className="empty-state">
                            <div className="empty-state-icon">
                                <span className="empty-state-emoji">📁</span>
                            </div>
                            <h2 className="empty-state-title" data-testid="text-empty-title">Select a Folder</h2>
                            <p className="empty-state-description" data-testid="text-empty-desc">
                                Choose a folder from the sidebar to view its contents.
                                Documents are displayed only at the final folder level.
                            </p>
                        </div>
                    )}
                </div>
            </div>

            {/* <NewFolderPanel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                onSubmit={handleCreateFolder}
                metadataFields={metadataFields}
            /> */}
        </div>
    );
};

export default React.memo(Workspace);