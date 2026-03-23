/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { DefaultButton, PrimaryButton, PanelType, Panel, DialogType, TextField, TooltipHost, DirectionalHint } from '@fluentui/react';
import { ArrowUpload20Regular, FolderAdd20Regular, Add20Regular, Home20Regular, ChevronRight12Regular, MoreHorizontalRegular, ChevronRight24Regular, ChevronDown24Regular } from '@fluentui/react-icons';
import Sidebar from "../../common/component/Sidebar";
import { FolderNode } from "../../common/component/FolderTree";
import { buildBreadcrumbPath, buildFolderHierarchy, buildLibraryRootPath, checkButtons, checkExtension, fileTypeConfig, getAllDocuments, getOpenAppURL } from "../../common/commonfunction";
import ReusableDataTable from "../ResuableComponents/ReusableDataTable";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import { ILabel } from "../../../../Intrface/ILabel";
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';
import * as FluentIcons from "@fluentui/react-icons";
import { Icon } from '@fluentui/react';
import {
    Menu,
    MenuTrigger,
    MenuPopover,
    MenuList,
    MenuItem,
    Button,
    Badge,
    Input,
    Label,
    Field
} from "@fluentui/react-components";
import { getAllButtons } from "../../../../Services/Buttons";
import { IButtonsProps, IRolePermission } from "../../../../Intrface/IButtonInterface";
import { checkPermissions, commonPostMethod, getApprovalData, getArchiveData, getListData, hasFolderPermission, updateLibrary } from "../../../../Services/GeneralDocument";
import { getHistoryByID } from "../../../../Services/GeneralDocHistoryService";
import { format } from "date-fns";
import { getConfigActive } from "../../../../Services/ConfigService";
import { getDataByLibraryName } from "../../../../Services/MasTileService";
import IFrameDialog from "../../common/component/IFrameDialog";
import AdvancePermission from "../../common/component/AdvancePermission";
import PopupBox, { ConfirmationDialog } from "../../common/component/PopupBox";
import { FolderStructure } from "../../../../Services/FolderStructure";
import { isMember } from "../../../../DAL/Commonfile";
import ProjectEntryForm from "../../common/component/ProjectEntryForm";
import UploadFiles from "../../common/component/UploadFile";
import ApprovalFlow from "../../common/component/ApprovalFlow";

interface IWorkspaceProps {
    context: WebPartContext;
}
interface Folder {
    [key: string]: string | number | {} | null | undefined;
}

const Workspace: React.FunctionComponent<IWorkspaceProps> = ({ context }) => {
    const SiteURL = context.pageContext.web.absoluteUrl;
    const UserID = context.pageContext.legacyPageContext.userId;
    const UserEmailID = context.pageContext.user.email;
    const portalUrl = new URL(context.pageContext.web.absoluteUrl).origin;
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const { workspaceId } = useParams<{ workspaceId: string; }>();
    const navigate = useNavigate();
    const [selectedFolder, setSelectedFolder] = useState<any | null>(null);
    const [folders, setFolders] = useState<any>([]);
    const [tileData, setTileData] = useState<any | null>(null);
    const [files, setFiles] = useState<any[]>([]);
    const [buttons, setButtons] = useState<any[]>([]);
    const [itemId, setItemId] = useState<number>(0);
    const [message, setMessage] = useState<string>("");
    const [hideDialog, setHideDialog] = useState<boolean>(false);
    const [actionButton, setActionButton] = useState<React.ReactNode>(null);
    const [panelForm, setPanelForm] = useState<React.ReactNode>(null);
    const [panelTitle, setPanelTitle] = useState("");
    const [isOpenCommonPanel, setIsOpenCommonPanel] = useState(false);
    const [extension, setExtension] = useState("");
    const [fileName, setFileName] = useState("");
    const [fileNameErr, setFileNameErr] = useState("");
    const [panelSize, setPanelSize] = useState(PanelType.medium);
    const [alertMsg, setAlertMsg] = useState("");
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState<boolean>(false);
    const [comment, setComment] = useState("");
    const [serverRelativeUrl, setServerRelativeUrl] = useState("");
    const [hideDialogCheckOut, setHideDialogCheckOut] = useState<boolean>(false);
    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [shareURL, setShareURL] = useState("");
    const [iFrameDialogOpened, setIFrameDialogOpened] = useState(false);
    const [isShowCommnPopupBoxVisible, setIsShowCommnPopupBoxVisible] = useState<boolean>(false);
    const [isOpenFolderPanel, setIsOpenFolderPanel] = useState(false);
    const [folderNameErr, setFolderNameErr] = useState("");
    const [folderName, setFolderName] = useState("");
    const invalidCharsRegex = /["*:<>?/\\|]/;
    const [admin, setAdmin] = useState([]);
    const [isValidUser, setIsValidUser] = useState<boolean>(false);
    const [isCreateProjectPopupOpen, setIsCreateProjectPopupOpen] = useState(false);
    const [isOpenUploadPanel, setIsOpenUploadPanel] = useState(false);
    const [fileType, setFileType] = useState<string>("");
    const [formType, setFormType] = useState("EntryForm");
    const [tables, setTables] = useState("");
    const [viewListSetting, setViewListSetting] = useState("");
    const [deletedData, setDeletedData] = useState<any>([]);
    const [approvalData, setApprovalData] = useState<any>([]);
    const [archiveData, setArchiveData] = useState<any>([]);
    const [projectUpdateData, setProjectUpdateData] = useState<any>({});
    const [hasPermission, setHasPermission] = useState<boolean>(false);
    const [isRestrictedView, setIsRestrictedView] = useState(false);
    const [expandedFolders, setExpandedFolders] = useState<string[]>([]);
    useEffect(() => {
        fetchTileData();
        getAdmin();
    }, []);

    const fetchTileData = async () => {
        const sp = spfi().using(SPFx(context));
        const data = await sp.web.lists.getByTitle("DMS_Mas_Tile").select("*").items.getById(Number(workspaceId))();
        setTileData(data);
    };

    useEffect(() => {
        if (tileData) {
            fetchFolder();
            getDeletedData();
            getArchiveFile();
            getUserGroups();
        }
    }, [tileData]);

    useEffect(() => {
        getPendingApprovalData();
    }, [isOpenUploadPanel, tileData]);




    const fetchFolder = async () => {
        const sp = spfi().using(SPFx(context));

        const allFolders: any[] = [];

        const items = await sp.web.lists
            .getByTitle(tileData?.LibraryName)
            .items
            .select("*", "Id", "Title", "FileRef", "FileDirRef", "FSObjType")
            .filter("FSObjType eq 1")
            .top(5000);

        for await (const batch of items) {
            allFolders.push(...batch);
        }

        const rootPath = buildLibraryRootPath(context, tileData?.LibraryName);
        const folder = buildFolderHierarchy(allFolders, rootPath);
        const folderObj = {
            id: 0,
            name: tileData?.LibraryName,
            path: tileData?.LibraryName,
            children: [...folder]
        };
        setFolders([folderObj]);
        expandParentFolders(folderObj);
        setSelectedFolder(folderObj);
    };


    const getAdmin = async () => {
        const data = await getListData(`${SiteURL}/_api/web/lists/getbytitle('DMS_GroupName')/items?`, context);
        setAdmin(data.value.map((el: any) => (el.GroupNameId)));
        const isMembers = await isMember(context, "ProjectAdmin");
        setIsValidUser(isMembers.value.length > 0);
    };

    const getDeletedData = async () => {
        const deletedData = await getListData(`${SiteURL}/_api/web/lists/getbytitle('${tileData?.LibraryName}')/items?$filter=DeleteFlag eq 'Deleted' and Active eq 0`, context);
        setDeletedData(deletedData.value);
    };

    const handleFolderSelect = (folder: FolderNode) => {
        setFiles([]);
        setSelectedFolder(folder);
        expandParentFolders(folder);
    };

    const getPendingApprovalData = async () => {
        const pendingApprovalData = await getApprovalData(context, tileData.LibraryName, UserEmailID);
        setApprovalData(pendingApprovalData.value);
    };
    const getArchiveFile = async () => {
        const data = await getArchiveData(context, tileData?.LibraryName);
        setArchiveData(data.value || []);
    };

    const handleFolderAction = (action: string, folder: FolderNode) => {
        console.log('Folder action:', action, folder.name);

        // const folderPath = selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, "");
        switch (action) {
            case "FView":
                setProjectUpdateData(folder); setIsCreateProjectPopupOpen(true); setFormType("ViewForm");
                break;
            case "FEdit":
                setProjectUpdateData(folder); setIsCreateProjectPopupOpen(true); setFormType("EditForm");
                break;
            case "AdvancePermission":
                setItemId(Number(folder.id)); setIsPanelOpen(true);
                break;
            case "Share":
                setShareURL(`${SiteURL}/_layouts/15/sharedialog.aspx?listId=${tileData?.LibGuidName}&listItemId=${folder.id}&clientId=sharePoint&policyTip=0&folderColor=undefined&ma=0&fullScreenMode=true&itemName=${folder.name}&origin=${portalUrl}`);
                setIFrameDialogOpened(true);
                break;
        }
    };

    const getRestrictedUserData = async () => {

        let View: any = "viewListItems";
        let Edit: any = "editListItems";

        const canView = await hasFolderPermission(
            context,
            selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, ""),
            View
        );

        const canEdit = await hasFolderPermission(
            context,
            selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, ""),
            Edit
        );

        // Restricted View Logic
        if (canView === true && canEdit === false) {
            setIsRestrictedView(true);
        } else {
            setIsRestrictedView(false);
        }
        console.log("VIEW:", canView, "EDIT:", canEdit, "Restricted:", canView && !canEdit);
    };


    const getPreviewUrl = (filePath: string) => {
        const extension = filePath?.split('.').pop()?.toLowerCase();
        switch (extension) {
            case 'doc':
            case 'docx':
            case 'ppt':
            case 'pptx':
            case 'xls':
            case 'xlsx':
                return <iframe src={`${SiteURL}/_layouts/15/WopiFrame.aspx?sourcedoc=${filePath}&action=embedview`} style={{ width: "100%", height: "80vh" }}></iframe>;

            case 'txt':
                return <iframe src={`${filePath}`} style={{ width: "100%", height: "80vh" }}></iframe>;
            case 'jpg':
            case 'jpeg':
            case 'png':
            case 'gif':

            case 'bmp':
                return <img src={`${filePath}`} alt={DisplayLabel.Preview} />;
            case 'pdf':
                return <iframe src={`${filePath}`} style={{ width: "100%", height: "80vh" }}></iframe>;
        }
    };

    const handleDocumentAction = async (action: string, item: any) => {
        switch (action) {
            case "OpenInApp":
                getOpenAppURL(item.ServerRelativeUrl, SiteURL);
                break;
            case "Delete":
                setMessage(DisplayLabel.DeleteConfirmMsg);
                setItemId(item.ListItemAllFields.Id);
                setHideDialog(true);
                break;
            case "Versions":
                setActionButton(null);
                const url = `${SiteURL}/_layouts/15/Versions.aspx?list=${tileData?.LibraryName}&FileName=${item.ServerRelativeUrl}&IsDlg=${item.ListItemAllFields.Id}`;
                setPanelForm(<iframe id="frame" src={url} style={{ width: "100%", height: "80vh" }}></iframe>);
                setPanelTitle(DisplayLabel.Versions);
                setIsOpenCommonPanel(true);
                break;
            case "Rename":
                setFileNameErr("");
                setItemId(item.ListItemAllFields.Id);
                setPanelTitle(DisplayLabel.Rename);
                const fileDetails = item.ListItemAllFields.ActualName.split(".");
                setExtension(fileDetails[1]);
                setFileName(fileDetails[0]);
                setIsOpenCommonPanel(true);
                break;
            case "Download":
                location.href = `${SiteURL}/_layouts/15/download.aspx?SourceUrl=${item.ServerRelativeUrl}`;
                break;
            case "Preview":
                setActionButton(null);
                setPanelSize(PanelType.smallFluid);
                setPanelTitle(DisplayLabel.Preview);
                const previewData = getPreviewUrl(item._original.ServerRelativeUrl);
                setPanelForm(previewData);
                setIsOpenCommonPanel(true);
                break;
            case "CheckOut":
                await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${item.ServerRelativeUrl}')/checkout`, context);
                setAlertMsg(DisplayLabel.CheckoutSuccess);
                setIsPopupBoxVisible(true);
                getDocument();
                break;
            case "CheckIn":
                setActionButton(<PrimaryButton text={DisplayLabel.CheckIn} style={{ marginRight: "10px" }} onClick={async () => {
                    await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${item.ServerRelativeUrl}')/checkin(comment='${comment}',checkintype=0)`, context);
                    setAlertMsg(DisplayLabel.CheckInSuccess);
                    setIsPopupBoxVisible(true);
                    getDocument();
                }} />);
                setIsOpenCommonPanel(true);
                break;
            case "DiscardCheckOut":
                setMessage(DisplayLabel.CheckoutConfirm);
                setServerRelativeUrl(item.ServerRelativeUrl);
                setHideDialogCheckOut(true);
                break;
            case "History":
                setActionButton(null);
                const HistoryData = await getHistoryByID(SiteURL, context.spHttpClient, item.ListItemAllFields.Id, tileData?.LibraryName);
                const bindData =
                    HistoryData?.value.length > 0 ? (
                        HistoryData.value
                            .sort((a: any, b: any) => {
                                return new Date(a.ActionDate).getTime() - new Date(b.ActionDate).getTime();
                            })
                            .map((el: any, index: number) => (
                                <tr key={index}>
                                    <td>{index + 1}</td>
                                    <td>{el.Action}</td>
                                    <td>{el.Author.Title}</td>
                                    <td>{el.ActionDate ? format(el.ActionDate, "DD-MM-YYYY hh:mm:ss A") : ""}</td>
                                    <td>{el.InternalComment}</td>
                                </tr>
                            ))
                    ) : (
                        <tr>
                            <td colSpan={5}>No Data</td>
                        </tr>
                    );
                setPanelForm(<table className="addoption" style={{ width: '100%', marginTop: '20px', borderCollapse: 'collapse' }}>
                    <thead>
                        <tr>
                            <th>{DisplayLabel?.SrNo}</th>
                            <th>{DisplayLabel?.Action}</th>
                            <th>{DisplayLabel?.ActionBy}</th>
                            <th>{DisplayLabel?.ActionDate}</th>
                            <th>{DisplayLabel?.Comments}</th>
                        </tr>
                    </thead>
                    <tbody>{bindData}</tbody>
                </table>);
                setPanelTitle(DisplayLabel.History);
                setIsOpenCommonPanel(true);
                break;
            case "View":
                setActionButton(null);
                const dataConfig = await getConfigActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
                const libraryData = await getDataByLibraryName(context.pageContext.web.absoluteUrl, context.spHttpClient, tileData.LibraryName);
                let jsonData = JSON.parse(libraryData.value[0].DynamicControl);
                jsonData = jsonData.filter((ele: any) => ele.IsActiveControl);
                //setPanelSize(PanelType.large);
                const htm = <>

                </>;
                setPanelForm(htm);
                setPanelTitle(DisplayLabel.View);
                setIsOpenCommonPanel(true);
                break;
            case "AdvancePermission":
                setItemId(item.ListItemAllFields.Id);
                setIsPanelOpen(true);
                break;
            case "Share":
                setShareURL(`${SiteURL}/_layouts/15/sharedialog.aspx?listId=${tileData.LibGuidName}&listItemId=${item.ListItemAllFields.Id}&clientId=sharePoint&policyTip=0&folderColor=undefined&ma=0&fullScreenMode=true&itemName=${item.ListItemAllFields.ActualName}&origin=${portalUrl}`);
                setIFrameDialogOpened(true);
                break;
            case "OpenInBrowser":
                const urls = item.LinkingUri === null ? item.ServerRelativeUrl : item.LinkingUri;
                window.open(urls, '_blank');
                break;
        }
    };

    const folderPathBread = useMemo<FolderNode[]>(() => {
        if (!selectedFolder) return [];
        return buildBreadcrumbPath(selectedFolder, folders);
    }, [selectedFolder]);

    useEffect(() => {
        if (selectedFolder) {
            getDocument();
            hasRequiredPermissions();
            getRestrictedUserData();
        }
    }, [selectedFolder]);

    useEffect(() => {
        getDocument();
    }, [isOpenUploadPanel]);

    const getDocument = async () => {
        if (!selectedFolder) return [];
        if (selectedFolder.isLastLevel) {
            const files = await getAllDocuments(context, selectedFolder?.path);
            setFiles(files.filter((el: any) => (el.ListItemAllFields.Active && (el.ListItemAllFields.InternalStatus === "Published" || el.ListItemAllFields.AuthorId === UserID))) || []);
        };
    };



    const getUserGroups = async () => {
        let allGroupCurrentUser: number[] = [];
        let filterData: any[] = [];

        try {

            const groupsResponse: SPHttpClientResponse = await context.spHttpClient.get(`${SiteURL}/_api/web/GetUserById(${UserID})/Groups`, SPHttpClient.configurations.v1);

            const groupsData = await groupsResponse.json();
            allGroupCurrentUser = groupsData.value.map((g: any) => g.Id);

            const buttonsResponse = await getAllButtons(SiteURL, context.spHttpClient);
            const allButtons: IButtonsProps[] = await buttonsResponse.value;


            const data: IRolePermission[] = JSON.parse(tileData?.CustomPermission);

            for (let i = 0; i < data.length; i++) {
                const filterData1 = data[i].UsersId.filter((u: any) => u.Id === UserID);
                if (filterData1.length > 0) {
                    filterData = filterData.concat(data[i]);
                }

                const filterData2 = data[i].UsersId.filter((u: any) => allGroupCurrentUser.includes(Number(u)));
                if (filterData2.length > 0) {
                    filterData = filterData.concat(data[i]);
                }
            }

            let filterData3: any[] = [];

            if (tileData.TileAdminId === UserID) {
                filterData3 = allButtons.map((btn) => ({
                    ...btn,
                    key: btn.InternalName
                }));
            } else {
                filterData.forEach((el) => {
                    allButtons.forEach((btn) => {
                        const match = el.Permission.filter((perm: any) => perm.value && btn.Id === perm.Id);
                        filterData3 = filterData3.concat([{ ...match[0], Icons: btn.Icons }]);
                    });
                });
            }

            const unique = filterData3.filter((el, index, self) => index === self.findIndex((p) => p.Title === el.Title));
            setButtons(unique);
            console.log('All Buttons:', unique);
        } catch (err) {
            console.error('Error in getUserGroups:', err);
        }
    };


    const createMenuProps = (item: any) => {
        return buttons.filter((btn) => btn.ButtonType === "Document")
            .filter((btn) => {
                switch (btn.key) {
                    case "OpenInApp":
                        const isCheck = checkExtension(item.data.Name);
                        return isCheck;
                    case "CheckIn":
                        return item.data.CheckOutType === 0 && item.data.CheckedOutByUser?.Id === UserID;
                    case "DiscardCheckOut":
                        return item.data.CheckOutType === 0 && item.data.CheckedOutByUser?.Id === UserID;
                    case "CheckOut":
                        return item.data.CheckOutType === 2;
                    case "Preview":
                        return !checkExtension(item.data.Name);
                    default:
                        return checkButtons(btn.key);
                }
            })
            .map((btn) => ({
                key: btn.key,
                text: btn.ButtonDisplayName,
                Icons: btn?.Icons
            }));
    };


    const getStatusStyles = (status: any) => {
        switch (status) {
            case "Pending With Approver":
                return { backgroundColor: "#f1faff", color: "#009ef7" };
            case "Published":
                return { backgroundColor: "#e8fff3", color: "#50cd89" };
            case "Pending With Publisher":
                return { backgroundColor: "#fff8dd", color: "#ffc700" };
            case "Rejected":
                return { backgroundColor: "#fff5f8", color: "#ed1c24" };
        }
    };

    const renderDocName = (item: any) => {
        const ext = item.Name.split(".").pop();
        const config = fileTypeConfig[ext] || fileTypeConfig.other;
        const { IconName, className } = config;
        const checkedOutUser = item?.CheckedOutByUser;
        const isCheckedOut = item?.CheckOutType === 0;
        const isCheckedOutByCurrentUser = checkedOutUser?.Id === UserID;
        return (
            <div className="doc-name-cell" data-testid={`link-document-${item.id}`}>
                <div className={`doc-icon-wrap ${className}`}>
                    <IconName className="doc-icon-svg" />
                </div>
                <span
                    className="table-cell-link"
                    onClick={() => {

                        if (item.LinkingUrl === "") {
                            if (isRestrictedView === true) {
                                const filePath = item.ServerRelativeUrl;
                                const folderPath = filePath.substring(0, filePath.lastIndexOf("/"));

                                const previewUrl = `${SiteURL}/${tileData?.LibraryName}/Forms/AllItems.aspx?id=${encodeURIComponent(filePath)
                                    }&parent=${encodeURIComponent(folderPath)}`;

                                window.open(previewUrl, "_blank");
                                return;
                            } else {
                                window.open(item.ServerRelativeUrl, "_blank");
                            }
                        }
                        else
                            window.open(item.LinkingUrl, "_blank");
                    }}
                >
                    {item?.ListItemAllFields?.ActualName}
                </span>
                {isCheckedOut && (
                    <TooltipHost
                        content={`${checkedOutUser?.Title} ${DisplayLabel.CheckedOutThisItem}`}
                        directionalHint={DirectionalHint.rightCenter} // Positioning
                        styles={{
                            root: { display: 'inline-block', maxWidth: '150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
                        }}
                    >
                        <Icon
                            iconName={isCheckedOutByCurrentUser ? "CheckedOutByYou12" : "CheckedOutByOther12"}
                            style={{ marginLeft: "5px", marginTop: '5px', color: isCheckedOutByCurrentUser ? "#a4262c" : "#605e5c", cursor: "pointer" }}
                        />

                    </TooltipHost>
                )}
            </div>
        );
    };

    const dismissCommanPanel = () => { setIsOpenCommonPanel(false); setActionButton(null); setPanelForm(null); setPanelSize(PanelType.medium); };
    const onDismiss: any = useCallback(() => { setIsPanelOpen(false); }, []);
    const closeDialog = useCallback(() => setHideDialog(false), []);
    const closeDialogCheckOut = useCallback(() => setHideDialogCheckOut(false), []);
    const hidePopup = useCallback(() => { setIsPopupBoxVisible(false); }, [isPopupBoxVisible]);
    const hideCommonPopup = useCallback(() => { setIsShowCommnPopupBoxVisible(false); }, []);
    const dismissFolderPanel = () => { setIsOpenFolderPanel(false); };
    const dissmissProjectCreationPanel = useCallback((value: boolean) => { setIsCreateProjectPopupOpen(value); fetchFolder(); }, []);
    const dismissUploadPanel = useCallback(() => { setIsOpenUploadPanel(false); }, []);

    const handleConfirm = useCallback(
        async (value: boolean) => {
            if (value) {
                setHideDialog(false);
                setIsPanelOpen(true);
                deleteDoc();
            }
        },
        [itemId]
    );
    const handleConfirmCheckOut = useCallback(async (value: boolean) => {
        if (value) {
            await commonPostMethod(`${SiteURL}/_api/web/GetFileByServerRelativeUrl('${serverRelativeUrl}')/undocheckout()`, context);
            setAlertMsg(DisplayLabel.DiscardedCheckOut);
            setIsPopupBoxVisible(true);
            getDocument();
        }
    }, [serverRelativeUrl]);

    const deleteDoc = async () => {
        const obj = {
            Active: false,
            DeleteFlag: "Deleted",
        };
        await updateLibrary(SiteURL, context.spHttpClient, obj, itemId, tileData.LibraryName);
        setAlertMsg(DisplayLabel.DeletedMsg);
        setIsPopupBoxVisible(true);
        getDocument();
    };

    useEffect(() => {
        setPanelForm(<>
            <div className="col-md-10">
                <Input value={fileName} required onChange={(_, val) => {
                    setFileName(val.value);
                }} />
            </div>
            <Label style={{ color: "red" }}>{fileNameErr}</Label>
            <div className="col-md-2"><Input readOnly value={extension} /></div>
        </>);
        setActionButton(<PrimaryButton text={DisplayLabel.Rename} style={{ marginRight: "10px" }} onClick={() => renameTheFile(itemId)} />);
    }, [fileName, extension, fileNameErr]);


    useEffect(() => {
        setPanelForm(<>
            <div className="col-md-10">
                <Input value={comment} onChange={(_, val) => setComment(val.value)} />
            </div>
        </>);

    }, [comment]);

    const renameTheFile = (id: number) => {
        if (fileName === "") {
            setFileNameErr(DisplayLabel.ThisFieldisRequired);
        }
        else {
            const obj = {
                ActualName: `${fileName}.${extension}`
            };
            updateLibrary(SiteURL, context.spHttpClient, obj, id, tileData.LibraryName).then((response) => {

                if (response) {
                    dismissFolderPanel();
                    setAlertMsg(DisplayLabel.SubmitMsg);
                    setIsPopupBoxVisible(true);
                    getDocument();
                }
                else {
                    dismissFolderPanel();
                    setAlertMsg(DisplayLabel.RenameAlertMsg);
                    setIsShowCommnPopupBoxVisible(true);
                }

            });
        }
    };

    const createFolder = (): void => {
        setFolderNameErr("");

        if (folderName === "") {
            setFolderNameErr(DisplayLabel.ThisFieldisRequired);
            return;
        }
        if (invalidCharsRegex.test(folderName)) {
            setFolderNameErr(DisplayLabel.FolderSpecialCharacterValidation);
            return;
        }
        if (!selectedFolder) return;

        const isDuplicate = selectedFolder.children.filter((el: any) => el.Name === folderName);
        if (isDuplicate.length > 0) {
            setFolderNameErr(DisplayLabel.FolderAlreadyExist);
            return;
        }

        const users = [selectedFolder?.ProjectmanagerId, selectedFolder?.PublisherId, ...admin];
        const siteRelative = context.pageContext.web.serverRelativeUrl;


        const urlAfterSite = selectedFolder.path.replace(siteRelative, "").replace(/^\/+/, "");

        FolderStructure(context, `${urlAfterSite}/${folderName}`, users, tileData.LibraryName, tileData.AllowChildInheritance).then(async (response) => {
            const sp = spfi().using(SPFx(context));
            const folderMetadata = await sp.web
                .getFolderByServerRelativePath(selectedFolder?.path)
                .listItemAllFields();
            const folderData = JSON.parse(JSON.stringify(folderMetadata, (key, value) => (value === null || (Array.isArray(value) && value.length === 0)) ? undefined : value));
            let obj: any = {
                ...folderData
            };

            updateLibrary(SiteURL, context.spHttpClient, obj, response, tileData.LibraryName).then((response) => {
                dismissFolderPanel();
                setAlertMsg(DisplayLabel.SubmitMsg);
                setIsPopupBoxVisible(true);
                fetchFolder();
            });
        });
    };


    const columns = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo,
                filter: false,
                resizable: false,
                maxWidth: 80,
                valueGetter: (params: any) => params.node.rowIndex + 1
            },
            {
                headerName: DisplayLabel.FileName,
                filter: true,
                sortable: true,
                field: "Name",
                maxWidth: 400,
                minWidth: 400,
                cellRenderer: (item: any) => renderDocName(item.data)
            },
            {
                headerName: DisplayLabel.ReferenceNo,
                filter: true,
                sortable: true,
                field: "ListItemAllFields.ReferenceNo",
                maxWidth: 160,
                minWidth: 120,
            },
            {
                headerName: DisplayLabel.Versions,
                filter: true,
                sortable: true,
                field: "ListItemAllFields.Level",
                maxWidth: 80,
                cellRenderer: (item: any) =>
                    <span className="table-cell-text table-cell-version" data-testid={`text-version-${item.id}`}>
                        v{item?.data?.ListItemAllFields?.Level}
                    </span>
            },
            {
                headerName: DisplayLabel.Status,
                filter: true,
                sortable: true,
                field: "ListItemAllFields.DisplayStatus",
                cellRenderer: (item: any) => {
                    const style = getStatusStyles(item.data.ListItemAllFields.DisplayStatus);
                    return <div>
                        <Badge
                            style={{ ...style }}
                        >
                            {item.data.ListItemAllFields.DisplayStatus}
                        </Badge>
                    </div>;
                }
            },
            {
                headerName: DisplayLabel.Action,
                filter: true,
                sortable: true,
                minWidth: 100,
                maxWidth: 120,
                cellRenderer: (item: any) => {
                    const menuProps = createMenuProps(item);
                    return <Menu>
                        <MenuTrigger disableButtonEnhancement>
                            <Button
                                appearance="subtle"
                                icon={<MoreHorizontalRegular className="table-action-btn" />}
                            />
                        </MenuTrigger>

                        <MenuPopover
                            style={{
                                boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                                padding: "15px"
                            }}
                        >
                            <MenuList>
                                {menuProps.map((e) => {
                                    // const IconComponent = FluentIcons[e.Icons as keyof typeof FluentIcons] as React.FC ?? <ChevronRight24Regular />;
                                    const IconComponent = (
                                        FluentIcons[e.Icons as keyof typeof FluentIcons] ??
                                        ChevronRight24Regular
                                    ) as React.ComponentType<React.SVGProps<SVGSVGElement>>;
                                    return <MenuItem
                                        key={e.key}
                                        icon={<IconComponent className="table-action-btn" />}
                                        onClick={() => handleDocumentAction(e.key, item?.data)}
                                    >
                                        {e.text}
                                    </MenuItem>;
                                })}
                            </MenuList>
                        </MenuPopover>
                    </Menu>;
                }
            }
        ];
    }, [buttons]);


    const renderRightFolder = (nodes: Folder[]) => {
        return (
            <div className="folder-grid">
                {nodes.map((node: any) => (
                    <div
                        key={node?.id}
                        className="folder-card"
                        onClick={() => handleFolderSelect(node)}
                    >
                        <FluentIcons.Folder20Filled className="folder-icon" />
                        <span className="folder-name">{node?.name}</span>
                    </div>
                ))}
            </div>
        );
    };

    const expandParentFolders = (folder: any) => {
        setExpandedFolders(prev => {
            if (prev.includes(folder?.id)) {
                return prev.filter(id => id !== folder?.id);
            } else {
                return [...prev, folder?.id];
            }
        });
    };

    const foldersColumn = React.useMemo(() => {
        return [
            {
                headerName: DisplayLabel.SrNo,
                filter: false,
                resizable: false,
                maxWidth: 80,
                valueGetter: (params: any) => params.node.rowIndex + 1
            },
            {
                headerName: DisplayLabel.FileName,
                filter: true,
                sortable: true,
                field: "Name",
                maxWidth: 400,
                minWidth: 400,
                cellRenderer: (item: any) => <a href="javascript:void()" onClick={() => handleFolderSelect(item?.data)} style={{ color: "rgb(0, 158, 247)" }}>{item?.data?.name}</a>
            },
            {
                headerName: DisplayLabel.LastModified,
                filter: false,
                resizable: false,
                maxWidth: 80,
                valueGetter: (params: any) => format(params?.data?.Modified, "dd-MM-yyyy hh:mm a")
            },
        ];
    }, []);

    const getItemStyle = (type: string) => ({
        display: "flex",
        alignItems: "center",
        gap: "10px",
        padding: "8px 12px",
        borderRadius: "6px",
        cursor: "pointer",
        backgroundColor: viewListSetting === type ? "#EAF3FC" : "transparent",
        color: viewListSetting === type ? "#0F6CBD" : "#323130",
        fontWeight: viewListSetting === type ? 600 : 400,
        transition: "all 0.2s ease"
    });

    const projectCreation = useCallback(() => { setIsCreateProjectPopupOpen(true); setFormType("EntryForm"); setProjectUpdateData({}); }, []);
    const hasRequiredPermissions = () => {
        checkPermissions(context, selectedFolder?.path).then((permission: boolean) => setHasPermission(permission));
    };
    const bindTable = () => {

        if (tables === "Approver") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Approver" />;
        }
        else if (tables === "Recycle") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Recycle" />;
        }
        else if (tables === "Archive") {
            return <ApprovalFlow context={context} libraryName={tileData?.LibraryName} userEmail={UserEmailID} action="Archive" />;
        }
        else {
            return (selectedFolder?.children.length === 0 && selectedFolder?.name !== tileData?.LibraryName) ?
                <ReusableDataTable rowData={files} columnDefs={columns} />
                :
                <div>
                    {viewListSetting === "List View" ? (
                        <ReusableDataTable rowData={selectedFolder?.children} columnDefs={foldersColumn} />
                    ) : (
                        <div >
                            {renderRightFolder(selectedFolder?.children)}
                        </div>
                    )}
                </div>;

        }

    };



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
                        onClick={projectCreation}
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
                    recycleBinCount={deletedData.length}
                    approvalCount={approvalData.length}
                    onRecycleBinClick={() => setTables("Recycle")}
                    onApprovalClick={() => navigate('/approvals', { state: { from: `/workspace/${workspaceId}`, libName: tileData?.LibraryName, tileName: tileData?.TileName } })}
                    onAdvancedSearchClick={() => navigate('/Search', { state: { from: `/workspace/${workspaceId}`, libName: tileData?.LibraryName } })}
                    onArchiveClick={() => { setTables("Archive"); }}
                    LibDetails={tileData}
                    archiveCount={archiveData.length}
                    buttons={buttons.filter((btn) => btn.ButtonType === "Folder")}
                    expandedFolders={expandedFolders}
                />

                <div className="workspace-content">
                    {selectedFolder && folderPathBread.length > 0 && (
                        <div className="workspace-content-header">
                            <div className="workspace-folder-breadcrumb" data-testid="nav-folder-breadcrumb">
                                {folderPathBread.map((node, i) => (
                                    <span key={node.id} className="workspace-folder-breadcrumb-segment">
                                        {i > 0 && <ChevronRight12Regular className="workspace-folder-breadcrumb-chevron" />}
                                        <span
                                            className={`workspace-folder-breadcrumb-item ${i === folderPathBread.length - 1 ? 'workspace-folder-breadcrumb-current' : ''}`}
                                            onClick={() => {
                                                if (i < folderPathBread.length - 1) handleFolderSelect(node);
                                            }}
                                            data-testid={`breadcrumb-folder-${node.id}`}
                                        >
                                            {node.name}
                                        </span>
                                    </span>
                                ))}
                            </div>
                            <div className="workspace-content-actions">
                                {tables === "" ? <>
                                    {(selectedFolder?.children.length === 0 && selectedFolder?.name !== tileData?.LibraryName) && (isValidUser || tileData?.TileAdminId === UserID || hasPermission) ?
                                        <Menu>
                                            <MenuTrigger disableButtonEnhancement>
                                                <Button
                                                    appearance="subtle"
                                                    iconPosition="after"
                                                    icon={<ChevronDown24Regular className="table-action-btn" />}
                                                    className="workspace-upload-btn"
                                                ><span>Create or Upload</span></Button>
                                            </MenuTrigger>

                                            <MenuPopover
                                                style={{
                                                    boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                                                    padding: "15px"
                                                }}
                                            >
                                                <MenuList>
                                                    <MenuItem
                                                        key="folder"
                                                        icon={<ArrowUpload20Regular style={{ color: "#0078D4" }} />}
                                                        onClick={() => {
                                                            setFileType("upload");
                                                            setIsOpenUploadPanel(true);
                                                        }}
                                                    >
                                                        Files Upload
                                                    </MenuItem>
                                                    <MenuItem
                                                        key="word"
                                                        icon={<Icon iconName="WordDocument" style={{ color: "#2B579A", fontSize: 20 }} />}
                                                        onClick={() => {
                                                            setFileType("docx");
                                                            setIsOpenUploadPanel(true);
                                                        }}
                                                    >
                                                        Word Document
                                                    </MenuItem>
                                                    <MenuItem
                                                        key="excel"
                                                        icon={<Icon iconName="ExcelDocument" style={{ color: "#217346", fontSize: 20 }} />}

                                                        onClick={() => {
                                                            setFileType("xlsx");
                                                            setIsOpenUploadPanel(true);
                                                        }}
                                                    >
                                                        Excel Document
                                                    </MenuItem>
                                                </MenuList>
                                            </MenuPopover>
                                        </Menu>
                                        // <DefaultButton text="Create or Upload" menuProps={uploadMenuProps} styles={{ root: { marginRight: 8 } }} />



                                        : <></>}
                                    {files.length === 0 && (hasPermission) ?
                                        <PrimaryButton
                                            onClick={() => { setIsOpenFolderPanel(true); setFolderName(""); setFolderNameErr(""); }}
                                            className="workspace-new-folder-btn"
                                            data-testid="button-new-folder"
                                        >
                                            <FolderAdd20Regular className="workspace-btn-icon" />
                                            <span>{DisplayLabel.NewFolder} </span>
                                        </PrimaryButton> : <></>}
                                </> : <> </>
                                }


                                {selectedFolder?.children.length !== 0 && (
                                    <Menu>
                                        <MenuTrigger disableButtonEnhancement>
                                            <Button
                                                appearance="transparent"
                                                iconPosition="after"
                                                icon={<FluentIcons.Board24Regular />}
                                            />
                                        </MenuTrigger>

                                        <MenuPopover
                                            style={{
                                                padding: "8px",
                                                borderRadius: "8px",
                                                boxShadow: "0 8px 24px rgba(0,0,0,0.15)",
                                                minWidth: "140px"
                                            }}
                                        >
                                            <div
                                                style={getItemStyle('List View')}
                                                onClick={() => setViewListSetting('List View')}
                                                onMouseEnter={(e) => {
                                                    if (viewListSetting !== 'List View')
                                                        e.currentTarget.style.backgroundColor = "#F3F2F1";
                                                }}
                                                onMouseLeave={(e) => {
                                                    if (viewListSetting !== 'List View')
                                                        e.currentTarget.style.backgroundColor = "transparent";
                                                }}
                                            >
                                                <FluentIcons.List20Regular
                                                    style={{
                                                        color: viewListSetting === 'List View' ? "#0F6CBD" : "#605E5C"
                                                    }}
                                                />
                                                List
                                            </div>
                                            <div
                                                style={getItemStyle('Tiles View')}
                                                onClick={() => setViewListSetting('Tiles View')}
                                                onMouseEnter={(e) => {
                                                    if (viewListSetting !== 'Tiles View')
                                                        e.currentTarget.style.backgroundColor = "#F3F2F1";
                                                }}
                                                onMouseLeave={(e) => {
                                                    if (viewListSetting !== 'Tiles View')
                                                        e.currentTarget.style.backgroundColor = "transparent";
                                                }}
                                            >
                                                <FluentIcons.Grid20Regular
                                                    style={{
                                                        color: viewListSetting === 'Tiles View' ? "#0F6CBD" : "#605E5C"
                                                    }}
                                                />
                                                Tiles
                                            </div>
                                        </MenuPopover>
                                    </Menu>
                                )}

                            </div>
                        </div>
                    )}

                    {selectedFolder ? (
                        <>
                            {bindTable()}
                        </>
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
            <Panel
                headerText={panelTitle}
                isOpen={isOpenCommonPanel}
                onDismiss={dismissCommanPanel}
                closeButtonAriaLabel="Close"
                type={panelSize}
                onRenderFooterContent={() => <>{actionButton}<DefaultButton onClick={dismissCommanPanel} >Cancel</DefaultButton></>}
                isFooterAtBottom={true}
            >
                <div style={{ marginTop: "10px" }}>
                    <div className="grid">
                        <div className="row">
                            {panelForm}
                        </div>
                    </div>
                </div>
            </Panel>
            <Panel
                headerText={DisplayLabel.AddNewFolder}
                isOpen={isOpenFolderPanel}
                onDismiss={dismissFolderPanel}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                onRenderFooterContent={() => (<>
                    <PrimaryButton onClick={createFolder} styles={{ root: { marginRight: 8 } }}>{DisplayLabel.Submit}</PrimaryButton>
                    <DefaultButton onClick={dismissFolderPanel}>{DisplayLabel.Cancel}</DefaultButton>
                </>)}
                isFooterAtBottom={true}
            >
                <Field>
                    <label>{DisplayLabel.Path}: <b>{
                        selectedFolder?.path
                            ?.replace(context.pageContext.web.serverRelativeUrl, "")
                            ?.replace(/^\/+/, "")
                    }</b></label>
                </Field>

                <Field >
                    <TextField
                        label={DisplayLabel.FolderName}
                        required value={folderName} onChange={(_, val) => {

                            setFolderName(val as string);

                            if (invalidCharsRegex.test(val as string)) {
                                setFolderNameErr(
                                    "Please enter a name that doesn't include any of these characters: \" * : < > ? / \\ |"
                                );
                            } else {
                                setFolderNameErr("");
                            }
                        }}
                        errorMessage={folderNameErr}
                    />
                    {/* <span style={{ color: "red" }}>{folderNameErr}</span> */}
                </Field>
            </Panel>

            <IFrameDialog
                url={shareURL}
                width="800px !important"
                height="600px"
                hidden={!iFrameDialogOpened}
                onDismiss={() => setIFrameDialogOpened(false)}
                iframeOnLoad={(iframe) => console.log('Iframe loaded:', iframe)}
                modalProps={{
                    isBlocking: true,

                }}
                dialogContentProps={{
                    type: DialogType.close,
                    showCloseButton: true
                }}
            />
            <AdvancePermission isOpen={isPanelOpen} context={context} folderId={itemId} LibraryName={tileData?.LibraryName} dismissPanel={onDismiss} />
            {tileData && <ProjectEntryForm isOpen={isCreateProjectPopupOpen} dismissPanel={dissmissProjectCreationPanel} context={context} LibraryDetails={tileData} admin={admin} FormType={formType} folderObject={projectUpdateData} folderPath={selectedFolder?.path} ChildFolderRoleInheritance={tileData?.AllowChildInheritance} />}
            <UploadFiles context={context} isOpenUploadPanel={isOpenUploadPanel} folderName={selectedFolder?.name} folderPath={selectedFolder?.path?.replace(context.pageContext.web.serverRelativeUrl, "")?.replace(/^\/+/, "")} dismissUploadPanel={dismissUploadPanel} libName={tileData?.LibraryName} files={files} folderObject={selectedFolder} LibraryDetails={tileData} filetype={fileType} FileData={files} />

            <ConfirmationDialog hideDialog={hideDialog} closeDialog={closeDialog} handleConfirm={handleConfirm} msg={message} />
            <ConfirmationDialog hideDialog={hideDialogCheckOut} closeDialog={closeDialogCheckOut} handleConfirm={handleConfirmCheckOut} msg={message} />
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} />
            <PopupBox isPopupBoxVisible={isShowCommnPopupBoxVisible} hidePopup={hideCommonPopup} msg={alertMsg} type="warning" />
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