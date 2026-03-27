import { WebPartContext } from "@microsoft/sp-webpart-base";
import React, { useCallback, useEffect, useState } from 'react';
import { getApprovalData, getRecycleData, getArchiveData, updateLibrary } from "../../../../Services/GeneralDocument";
import { DefaultButton, FontIcon, Panel, PanelType, PrimaryButton, TextField } from "@fluentui/react";
import { getStatusByInternalStatus } from "../../../../Services/StatusSerivce";
import { createHistoryItem } from "../../../../Services/GeneralDocHistoryService";
import { TileSendMail } from "../../../../Services/SendEmail";
import PopupBox, { ConfirmationDialog } from "./PopupBox";
import { ILabel } from "../../../../Intrface/ILabel";
import { format } from "date-fns";
import ReusableDataTable from "../../components/ResuableComponents/ReusableDataTable";

interface IApproval {
    context: WebPartContext;
    libraryName: string;


    userEmail: string;
    action: string;
}
const ApprovalFlow: React.FunctionComponent<IApproval> = ({ context, libraryName, userEmail, action }) => {
    const [files, setFiles] = useState([]);
    const buttonStyles = { root: { marginRight: 8 } };
    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [comment, setcomment] = useState("");
    const [commentErr, setcommentErr] = useState("");
    const [fileData, setFileData] = useState<any>([]);
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState<boolean>(false);
    const [popupType, setPopupType] = useState<"success" | "warning" | "insert" | "checkin" | "checkout" | "approve" | "reject" | "delete" | "update" | "restore" | "grant" | "remove">("success");
    const [itemId, setItemId] = useState(0);
    const [hideDialog, setHideDialog] = useState(false);
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const [alertMsg, setAlertMsg] = useState("");

    useEffect(() => {
        if (action === "Approver")
            getFiles();
        else if (action === "Archive")
            getArchiveFile();
        else
            getRecycleFile();

    }, [isPopupBoxVisible, action]);

    const getFiles = async () => {
        const data = await getApprovalData(context, libraryName, userEmail);
        setFiles(data.value || []);
    };
    const getRecycleFile = async () => {
        const data = await getRecycleData(context, libraryName);
        setFiles(data.value || []);
    };
    const getArchiveFile = async () => {
        const data = await getArchiveData(context, libraryName);
        setFiles(data.value || []);
    };
    const visibleColumns = [
        "Status.StatusName",
        "Id"
    ];

    // const truncateText = (text: string, maxLength: number) => {
    //     return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
    // };
    const columns: any = [
        {
            Header: DisplayLabel.FileName, accessor: "Name", Cell: ({ row }: { row: any; }) => <a style={{
                color: "#009ef7",
                display: "inline-block",
                inlineSize: "150px",
                overflowWrap: "break-word",
                wordBreak: "break-word",
                whiteSpace: "normal"
            }} href="javascript:void('0')" onClick={() => {
                if (row._original.File.LinkingUrl === "")
                    window.open(row._original.File.ServerRelativeUrl, "_blank");
                else
                    window.open(row._original.File.LinkingUrl, "_blank");
            }}> {row._original?.ActualName}</a>
        },
        { Header: DisplayLabel.FolderPath, accessor: 'FolderDocumentPath' },
        {
            Header: 'Submitted By', accessor: 'Created', Cell: ({ row }: { row: any; }) => {
                const rowData = row._original;
                const formattedDate = format(rowData.Created, "dd/MM/yyyy");
                // new Date(rowData.Created).toLocaleDateString("en-US", {
                //     month: "2-digit",
                //     day: "2-digit",
                //     year: "numeric"
                // });
                const formattedTime = new Date(rowData.Created).toLocaleTimeString("en-US", {
                    hour: "2-digit",
                    minute: "2-digit",
                    hour12: true
                });
                return `${rowData.Author?.Title || "Unknown"} ${formattedDate} at ${formattedTime}`;
            }
        },
        { Header: DisplayLabel.Status, accessor: 'Status.StatusName' },
        {
            Header: DisplayLabel.Action,
            accessor: 'Id',
            Cell: ({ row }: { row: any; }) => (
                action === "Approver" ? <FontIcon aria-label="Edit" onClick={() => openEditPanel(row._original.Id)} iconName="EditSolid12" style={{ color: '#009ef7', cursor: 'pointer' }}></FontIcon> :
                    <FontIcon aria-label="Restore" title="Restore" onClick={() => { setItemId(row._original.Id); setHideDialog(true); }} iconName="RemoveFromTrash" style={{ color: '#009ef7', cursor: 'pointer' }} />
            )
        }
    ].filter(column => action === "Archive" ? !visibleColumns.includes(column.accessor) : true);
    const closeDialog = useCallback(() => { setHideDialog(false); setcomment(""); }, []);

    const handleConfirm = useCallback(
        async (value: boolean) => {
            if (value) {
                setHideDialog(false);
                restoreFile();
            }
        },
        [itemId]
    );
    const restoreFile = async () => {
        const obj = {
            Active: true,
            DeleteFlag: false // Instead of null
        };
        await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, obj, itemId, libraryName);
        setAlertMsg(DisplayLabel.RestoreDoc);
        setPopupType("restore");
        setIsPopupBoxVisible(true);
    };

    const openEditPanel = async (rowData: any) => {
        setcomment("");
        setIsPanelOpen(true);
        const fData = files.find((el: any) => el.Id === rowData);
        setFileData(fData);
    };

    const ApproveFile = async () => {
        setcommentErr("");
        if (comment === null || comment.trim() === "" || comment === undefined)
            setcommentErr(DisplayLabel.ThisFieldisRequired);
        else {
            try {
                let dataObj: any = {};
                let InternalStatus = "", ToUser = "";
                if (fileData.InternalStatus === "PendingWithPM" && fileData.PublisherEmail !== null) {
                    dataObj.CurrentApprover = fileData.PublisherEmail;
                    ToUser = fileData.PublisherEmail;
                    InternalStatus = "PendingWithPublisher";
                } else {
                    dataObj.CurrentApprover = "";
                    let PMEmail = fileData.ProjectmanagerEmail;
                    let AuthorEmail = fileData.Author.EMail;
                    ToUser = (PMEmail == "" ? AuthorEmail : (PMEmail + ";" + AuthorEmail));
                    InternalStatus = "Published";
                }
                dataObj.LatestRemark = comment;
                const status = await getStatusByInternalStatus(context.pageContext.web.absoluteUrl, context.spHttpClient, InternalStatus);

                dataObj.StatusId = status.value[0].ID;
                dataObj.InternalStatus = status.value[0].InternalStatus;
                dataObj.DisplayStatus = status.value[0].StatusName;
                await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, dataObj, fileData.Id, libraryName);

                var dataHistory = {
                    DocumetLID: fileData.Id,
                    ActionDate: new Date(),
                    Remark: comment,
                    Status: status.value[0].StatusName,
                    InternalComment: comment,
                    LibName: libraryName,
                    Action: "Approved"
                };
                await createHistoryItem(context.pageContext.web.absoluteUrl, context.spHttpClient, dataHistory);
                var emailObj: any = {
                    To: ToUser,
                    FolderPath: fileData.FolderDocumentPath,
                    DocName: fileData.ActualName,
                    AuthorTitle: fileData.Author.Title,
                    TileName: libraryName
                };

                if (InternalStatus == "PendingWithPublisher") {
                    emailObj.Sub = DisplayLabel.PublisherEmailSubject + " " + fileData.ReferenceNo;
                    //emailObj.Msg = DisplayLabel.PublisherEmailMsg;
                    emailObj.Status = InternalStatus;
                } else if (InternalStatus == "PendingWithPM") {
                    emailObj.Sub = DisplayLabel.PMEmailSubject + " " + fileData.ReferenceNo;
                    //emailObj.Msg = DisplayLabel.PMEmailMsg;
                    emailObj.Status = InternalStatus;

                } else {
                    emailObj.Sub = DisplayLabel.PublishedEmailSubject + " " + fileData.ReferenceNo;
                    emailObj.Msg = DisplayLabel.PublishedEmailMsg;
                    emailObj.Status = InternalStatus;
                }
                emailObj.ID = fileData.Id;
                emailObj.libraryName = libraryName;
                await TileSendMail(context, emailObj);
                setAlertMsg(DisplayLabel.ApprovedMsg);
                setPopupType("approve");
                setIsPopupBoxVisible(true);
                // window.location.reload();

            } catch (error) {
                console.log("error", error);
            }
        }
    };

    const dismissPanel = () => {
        setIsPanelOpen(false);
    };
    const hidePopup = useCallback(() => {
        setIsPopupBoxVisible(false);
        setIsPanelOpen(false);
    }, [isPopupBoxVisible]);

    const RejectFile = async () => {
        setcommentErr("");
        if (comment === null || comment.trim() === "" || comment === undefined)
            setcommentErr(DisplayLabel.ThisFieldisRequired);
        else {
            let InternalStatus = "";
            let dataobj: any = { CurrentApprover: "" };
            InternalStatus = "Rejected";
            dataobj.LatestRemark = comment;
            let ToUser = fileData.Author.EMail;
            if (fileData.InternalStatus !== "PendingWithPM") {
                ToUser = (fileData.ProjectmanagerEmail === "" ? ToUser : (ToUser + ";" + fileData.ProjectmanagerEmail));
            }

            const status = await getStatusByInternalStatus(context.pageContext.web.absoluteUrl, context.spHttpClient, InternalStatus);

            dataobj.StatusId = status.value[0].ID;
            dataobj.InternalStatus = status.value[0].InternalStatus;
            dataobj.DisplayStatus = status.value[0].StatusName;

            await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, dataobj, fileData.Id, libraryName);

            var dataHistory = {
                DocumetLID: fileData.Id,
                ActionDate: new Date(),
                Remark: comment,
                Status: status.value[0].StatusName,
                InternalComment: comment,
                LibName: libraryName,
                Action: "Rejected"
            };

            await createHistoryItem(context.pageContext.web.absoluteUrl, context.spHttpClient, dataHistory);
            var emailObj = {
                To: ToUser,
                FolderPath: fileData.FolderDocumentPath,
                DocName: fileData.ActualName,
                AuthorTitle: fileData.Author.Title,
                TileName: libraryName,
                Sub: DisplayLabel.RejectEmailSubject + " " + fileData.ReferenceNo,
                Msg: DisplayLabel.RejectEmailMsg,
                Status: status.value[0].StatusName
            };

            await TileSendMail(context, emailObj);
            setAlertMsg(DisplayLabel.RejectedMsg);
            setPopupType("reject");
            setIsPopupBoxVisible(true);
        }
    };

    return (
        <>
            <ReusableDataTable rowData={files} columnDefs={columns} />

            <Panel
                headerText={DisplayLabel.Approval}
                isOpen={isPanelOpen}
                onDismiss={dismissPanel}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                onRenderFooterContent={() => (<div style={{ marginBottom: "10px" }}>
                    <PrimaryButton onClick={ApproveFile} styles={buttonStyles} >{DisplayLabel.ApproveButton}</PrimaryButton>
                    <DefaultButton onClick={RejectFile}>{DisplayLabel.RejectButton}</DefaultButton>
                </div>)}
                isFooterAtBottom={true}
            >
                <div >
                    <div className="row">
                        <div className="col-md-12">
                            <TextField
                                label={DisplayLabel.TileName}
                                value={libraryName}
                                readOnly
                            />
                        </div>
                    </div>

                    <div className="row">
                        <div className="col-md-12">
                            <TextField
                                label={DisplayLabel.FolderName}
                                value={fileData?.FolderDocumentPath
                                    ? fileData.FolderDocumentPath.split("/").pop() || ""
                                    : ""}
                                readOnly
                            />
                        </div>
                    </div>


                    <div className="row">
                        <div className="col-md-12">
                            <TextField
                                label={DisplayLabel.AttachFile}
                                value={fileData.ActualName}
                                readOnly
                            />
                        </div>
                    </div>

                    <div className="row">
                        <div className="col-md-12">
                            <TextField
                                label={DisplayLabel.Comments}
                                value={comment}
                                multiline rows={3}
                                required
                                errorMessage={commentErr}
                            // onChange={(e: React.ChangeEvent<HTMLInputElement>) => setcomment(e.target.value)}
                            />
                        </div>
                    </div>
                </div>

            </Panel>
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} type={popupType} />
            <ConfirmationDialog hideDialog={hideDialog} closeDialog={closeDialog} handleConfirm={handleConfirm} msg={DisplayLabel.RestoreConfirmMsg} />
        </>
    );
};
export default React.memo(ApprovalFlow);
