import { ChoiceGroup, DatePicker, DefaultButton, Dropdown, IconButton, mergeStyleSets, Panel, PanelType, PrimaryButton, TextField } from "@fluentui/react";
import React, { useCallback, useEffect, useState } from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { uuidv4 } from "../../../../DAL/Commonfile";
import { getConfigActive } from "../../../../Services/ConfigService";
import { generateAutoRefNumber, getListData, updateLibrary, UploadFile, getDataByRefID } from "../../../../Services/GeneralDocument";
import { getDataByLibraryName } from "../../../../Services/MasTileService";
import PopupBox, { ConfirmationDialog } from "./PopupBox";
import { getStatusByInternalStatus } from "../../../../Services/StatusSerivce";
import { ILabel } from "../../../../Intrface/ILabel";
import Select from "react-select";
import moment from "moment";
import { TileSendMail } from "../../../../Services/SendEmail";
import { Field } from "@fluentui/react-components";
import FieldError from "./FieldError";
import PageLoader from "./PageLoader";

interface IUploadFileProps {
    isOpenUploadPanel: boolean;
    dismissUploadPanel: () => void;
    folderPath: string;
    libName: string;
    folderName: string;
    context: WebPartContext;
    files: any;
    folderObject: any;
    LibraryDetails: any;
    filetype: string;
    FileData: any[];
}
function UploadFiles({ context, isOpenUploadPanel, dismissUploadPanel, folderPath, libName, folderName, files, folderObject, LibraryDetails, filetype, FileData }: IUploadFileProps) {
    // const fileInputRef = useRef<HTMLInputElement | null>(null);
    const inValidExtensions = ["exe", "mp4", "mp3"];
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const [configData, setConfigData] = useState<any[]>([]);
    const [dynamicControl, setDynamicControl] = useState<any[]>([]);
    // const [libraryDetails, setLibraryDetails] = useState<any>({});
    const [options, setOptions] = useState<any>({});
    const [dynamicValues, setDynamicValues] = useState<{ [key: string]: any; }>({});
    const [dynamicValuesErr, setDynamicValuesErr] = useState<{ [key: string]: string; }>({});
    const [attachmentsFiles, setAttachmentsFiles] = useState<any[]>([]);
    //    const [attachment, setAttachment] = useState<{ [key: string]: any; }>({});
    const [attachment, setAttachment] = React.useState<File[]>([]);
    const [attachmentErr, setAttachmentErr] = useState<string>('');
    const [filesData, setFilesData] = useState<any[]>([]);
    const [filterFilesData, setFilterFilesData] = useState<any[]>([]);
    const [existingFile, setExistingFile] = useState<any[]>([]);
    const [isUpdateExistingFile, setIsUpdateExistingFile] = useState<boolean>(false);
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState<boolean>(false);
    const [showLoader, setShowLoader] = useState({ display: "none" });
    const [fileKey, setFileKey] = useState<number>(Date.now());
    const [alertMsg, setAlertMsg] = useState("");
    const [archiveCount, setArchiveCount] = useState("");
    const [newFileName, setNewFileName] = useState("");
    const [fileNameError, setFileNameError] = useState("");
    const invalidCharsRegex = /["*:<>?/\\|]/;

    const [showConfirmDialog, setShowConfirmDialog] = useState(false);
    const [duplicateFiles, setDuplicateFiles] = useState<File[]>([]);
    const [existingFileNamesInFileData, SetexistingFileNamesInFileData] = useState<any[]>([]);




    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any
    };
    const meargestyles = mergeStyleSets({
        root: { selectors: { '> *': { marginBottom: 15 } } },
        control: { maxWidth: "100%", marginBottom: 15 },
    });



    useEffect(() => {
        fetchLibraryDetails();
    }, []);

    useEffect(() => {
        setAttachmentErr("");
        setDynamicValuesErr({});
        setDynamicValues({});
        setAttachmentsFiles([]);
        setExistingFile([]);
        setFilesData([]);
        setFilterFilesData([]);
    }, [isOpenUploadPanel]);

    const handleInputChange = (key: string, value: any) => {
        setDynamicValues((prev) => ({ ...prev, [key]: value }));
    };
    const fetchLibraryDetails = async () => {
        const dataConfig = await getConfigActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
        const libraryData = await getDataByLibraryName(context.pageContext.web.absoluteUrl, context.spHttpClient, libName);

        // setLibraryDetails(libraryData.value[0]);
        setConfigData(dataConfig.value);

        if (libraryData.value[0]?.DynamicControl) {
            let jsonData = JSON.parse(libraryData.value[0].DynamicControl);
            jsonData = jsonData.filter((ele: any) => ele.IsActiveControl && ele.IsFieldAllowInFile);
            jsonData = jsonData.map((el: any) => {
                if (el.ColumnType === "Person or Group") {
                    el.InternalTitleName = `${el.InternalTitleName}Id`;
                }
                return el;
            });
            setDynamicControl(jsonData);
            bindDropdown(jsonData);
        }

        if (libraryData.value[0]?.ArchiveVersionCount) {
            setArchiveCount(libraryData.value[0]?.ArchiveVersionCount);
        }
    };

    const bindDropdown = (dynamic: any) => {
        let dropdownOptions = [{ key: "", text: "" }];
        dynamic.map(async (item: any, index: number) => {
            if (item.ColumnType === "Dropdown" || item.ColumnType === "Multiple Select") {
                if (item.IsStaticValue) {
                    dropdownOptions = item.StaticDataObject.split(";").map((ele: string) => ({
                        value: ele,
                        label: ele,
                    }));
                } else {
                    const data = await getListData(`${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${item.InternalListName}')/items?$top=5000&$filter=Active eq 1&$orderby=${item.DisplayValue} asc`, context);
                    dropdownOptions = data.value.map((ele: any) => ({
                        value: ele[item.DisplayValue],
                        label: ele[item.DisplayValue],
                    }));
                }
                setOptions((prev: any) => ({ ...prev, [item.InternalTitleName]: dropdownOptions }));
            }
        });
    };

    const removeSepcialCharacters = (newValue?: string) => newValue?.replace(/[^a-zA-Z0-9\s]/g, '') || '';

    const renderDynamicControls = useCallback(() => {
        return dynamicControl.map((item: any, index: number) => {
            const filterObj = configData.find((ele) => ele.Id === item.Id);

            if (!filterObj) return null;

            switch (item.ColumnType) {
                case "Dropdown":
                case "Multiple Select":
                    return (
                        <div className="column6" key={index}>
                            <Field label={item.Title} required={item.IsRequired}>
                                <Select
                                    options={options[item.InternalTitleName]}
                                    required={item.IsRequired}
                                    value={(options[item.InternalTitleName] || []).find((option: any) => option.value === dynamicValues[item.InternalTitleName])}
                                    onChange={(option: any) => handleInputChange(item.InternalTitleName, option?.value)}
                                    isSearchable
                                    placeholder={DisplayLabel?.Selectanoption}
                                    isMulti={item.ColumnType === "Multiple Select"}
                                />
                                <FieldError message={dynamicValuesErr[item.InternalTitleName]} />
                            </Field>
                        </div>
                    );

                case "Person or Group":
                    return (
                        <div className="column6" key={index}>
                            <PeoplePicker
                                titleText={item.Title}
                                context={peoplePickerContext}
                                personSelectionLimit={20}
                                showtooltip={true}
                                required={item.IsRequired}
                                showHiddenInUI={false}
                                ensureUser={true}
                                principalTypes={[PrincipalType.User]}
                                onChange={(users: any[]) => {
                                    const userIds = users.map(user => user.id);
                                    setDynamicValues((prevValues) => ({
                                        ...prevValues,
                                        [item.InternalTitleName]: userIds
                                    }));

                                }}
                                errorMessage={dynamicValuesErr[item.InternalTitleName]}
                            />
                        </div>
                    );

                case "Radio":
                    const radioOptions = filterObj.StaticDataObject.split(";").map((ele: string) => ({
                        key: ele,
                        text: ele,
                    }));
                    return (
                        <div className="column6" key={index}>
                            <ChoiceGroup
                                options={radioOptions}
                                onChange={(ev, option) => handleInputChange(item.InternalTitleName, option?.key)}
                                selectedKey={dynamicValues[item.InternalTitleName] || ""}
                                label={item.Title}
                                required={item.IsRequired}
                            />
                        </div>
                    );

                case "Date and Time":
                    return (
                        <div className="column6" key={index}>

                            <Field label={item.Title} required={item.IsRequired}>
                                <DatePicker
                                    onSelectDate={(date: Date | null | undefined) => handleInputChange(item.InternalTitleName, date)}
                                    className={meargestyles.control}
                                    value={dynamicValues[item.InternalTitleName] || ""}
                                    formatDate={(date) => date ? moment(new Date(date)).format("DD/MM/YYYY") : ''}
                                />
                                <FieldError message={dynamicValuesErr[item.InternalTitleName]} />
                            </Field>
                        </div>
                    );

                default:
                    return (
                        <div className="column6" key={index}>

                            <Field label={item.Title}>

                                <TextField
                                    type={"text"}
                                    value={dynamicValues[item.InternalTitleName] || ""}
                                    onChange={(ev, value) => handleInputChange(item.InternalTitleName, removeSepcialCharacters(value))}
                                    multiline={item.ColumnType === "Multiple lines of Text"}
                                    required={item.IsRequired}
                                    errorMessage={dynamicValuesErr[item.InternalTitleName]}
                                />
                            </Field>
                        </div>
                    );
            }
        });
    }, [dynamicControl, options, dynamicValues, dynamicValuesErr]);

    const addAttachment = () => {
        if (!attachment.length) {
            setAttachmentErr(DisplayLabel.ThisFieldisRequired);
            return;
        }

        const validFiles: File[] = [];
        let duplicateFound = false;

        for (const file of attachment) {
            const ext = file.name.split('.').pop()?.toLowerCase();

            if (ext && inValidExtensions.includes(ext)) {
                setAttachmentErr(DisplayLabel.InvalidFileFormat);
                return;
            }

            const existsInFileData = FileData?.find(
                f => f.ListItemAllFields.ActualName?.toLowerCase() === file.name.toLowerCase()
            );

            const existsInAttachments = attachmentsFiles.some(
                att => att.attachment.name.toLowerCase() === file.name.toLowerCase()
            );

            if (existsInFileData || existsInAttachments) {
                duplicateFound = true;
                SetexistingFileNamesInFileData(prev => [...prev, file.name]);
                setDuplicateFiles(prev => [...prev, file]);
                continue;
            }


            validFiles.push(file);
        }

        if (duplicateFound) {
            setShowConfirmDialog(true);
        }

        if (duplicateFound) {
            setAttachmentErr("File with the same name already exists.");
        }
        if (!validFiles.length) return;

        const newAttachments = validFiles.map(file => ({
            attachment: file,
            isUpdateExistingFile: "No",
            OldFileName: "",
            version: "1.0",
            isDisabled: true,
            Flag: "New"
        }));

        setAttachmentsFiles(prev => [...prev, ...newAttachments]);
        setAttachment([]);
        setAttachmentErr("");
        setFileKey(Date.now()); // reset file input
    };


    const renameWithCounter = (file: File) => {
        const dotIndex = file.name.lastIndexOf(".");
        const name = file.name.substring(0, dotIndex);
        const ext = file.name.substring(dotIndex);

        let counter = 1;
        let newName = `${name}(${counter})${ext}`;

        const isExist = (fileName: string) =>
            FileData?.some(f =>
                f.ListItemAllFields?.ActualName?.toLowerCase() === fileName.toLowerCase()
            ) ||
            attachmentsFiles?.some(f =>
                f.attachment.name.toLowerCase() === fileName.toLowerCase()
            );

        while (isExist(newName)) {
            counter++;
            newName = `${name}(${counter})${ext}`;
        }

        return new File([file], newName, { type: file.type });
    };

    const handleDuplicateConfirm = async (keepBoth: boolean) => {
        setShowConfirmDialog(false);

        if (!duplicateFiles.length) return;

        if (keepBoth) {
            const renamedAttachments = duplicateFiles.map(file => {
                const renamedFile = renameWithCounter(file);
                return {
                    attachment: renamedFile,
                    isUpdateExistingFile: "No",
                    OldFileName: "",
                    version: "1.0",
                    isDisabled: true,
                    Flag: "New"
                };
            });

            setAttachmentsFiles(prev => [...prev, ...renamedAttachments]);
        }
        else {
            const replaceAttachments = duplicateFiles.map(file => ({
                attachment: file,
                isUpdateExistingFile: "No",
                OldFileName: "",
                version: "1.0",
                isDisabled: true,
                Flag: "Replace"
            }));

            setAttachmentsFiles(prev => [...prev, ...replaceAttachments]);
        }

        setDuplicateFiles([]);
        setFileKey(Date.now());
    };





    const onClickDetails = (index: number) => {
        let IsExistingReferenceNo = "";
        if (attachmentsFiles[index].isUpdateExistingFile === "Yes") {
            let eFile = filterFilesData.filter((ele: any) => ele.Name == attachmentsFiles[index].OldFileName);
            IsExistingReferenceNo = eFile.length > 0 ? eFile[0].ListItemAllFields.IsExistingRefID : "";
            setExistingFile((per) => [...per, { ...eFile[0] }]);
            if (attachmentsFiles[index].OldFileName === "" || attachmentsFiles[index].OldFileName === null) {
                setAttachmentErr(DisplayLabel.ThisFieldisRequired);
                return false;
            }
        }
        setAttachmentsFiles((prev) => prev.map((ele, i) => i === index ? { ...ele, isDisabled: !ele.isDisabled, IsExistingRefID: IsExistingReferenceNo } : ele));
    };


    const validateFileName = (): boolean => {
        let isFileValid = true;
        if (dynamicControl.length > 0) {
            dynamicControl.forEach((item: any) => {
                if (item.IsRequired && !dynamicValues[item.InternalTitleName]) {
                    setDynamicValuesErr((prev) => ({
                        ...prev,
                        [item.InternalTitleName]: DisplayLabel.ThisFieldisRequired,
                    }));
                    isFileValid = false;
                }
            });
        }

        const trimmedName = newFileName.trim();
        if (!trimmedName) {
            setFileNameError(DisplayLabel.ThisFieldisRequired);
            return false;
        }

        if (invalidCharsRegex.test(trimmedName)) {
            setFileNameError("File name contains invalid characters");
            return false;
        }

        const existsInFileData = FileData?.find(f => {
            const actualName = f.ListItemAllFields?.ActualName;
            if (!actualName) return false;

            const nameWithoutExt =
                actualName.substring(0, actualName.lastIndexOf(".")) || actualName;

            return nameWithoutExt.toLowerCase() === trimmedName.toLowerCase();
        });

        if (existsInFileData) {
            setFileNameError("File Name already exists");
            return false;
        }

        setFileNameError("");
        return isFileValid;

    };

    const getDigest = async () => {
        const response = await fetch(
            `${context.pageContext.web.absoluteUrl}/_api/contextinfo`,
            {
                method: "POST",
                headers: {
                    "Accept": "application/json;odata=verbose"
                }
            }
        );

        const data = await response.json();
        return data.d.GetContextWebInformation.FormDigestValue;
    };

    const createOfficeFile = async () => {

        if (!validateFileName()) return;

        try {
            const digest = await getDigest();

            const folderServerRelativeUrl =
                `${context.pageContext.web.serverRelativeUrl}/${folderPath}`
                    .replace(/\/+/g, "/");

            const fileName = `${newFileName}.${filetype}`;
            const fileServerRelativeUrl = `${folderServerRelativeUrl}/${fileName}`;

            console.log("Final Path:", folderServerRelativeUrl);

            const createRes = await fetch(
                `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${folderServerRelativeUrl}')/Files/add(overwrite=true,url='${fileName}')`,
                {
                    method: "POST",
                    body: new ArrayBuffer(0),
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-RequestDigest": digest
                    }
                }
            );

            if (!createRes.ok) {
                throw new Error("File creation failed");
            }
            const itemResponse = await fetch(
                `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${fileServerRelativeUrl}')/ListItemAllFields`,
                {
                    headers: {
                        "Accept": "application/json;odata=verbose"
                    }
                }
            );

            const itemData = await itemResponse.json();
            const itemId = itemData.d.Id;

            console.log("Created Item ID:", itemId);


            const res = await fetch(
                `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libName}')?$select=ListItemEntityTypeFullName`,
                { headers: { Accept: "application/json;odata=verbose" } }
            );

            const data = await res.json();
            const entityType = data.d.ListItemEntityTypeFullName;
            console.log(entityType);


            let obj: any = {
                ActualName: fileName,
                FolderDocumentPath: `/${folderPath}`,
                OCRStatus: "Pending",
                UploadFlag: "Frontend",
                Level: "1.0",
                Active: true
            };

            let InternalStatus = "Published";

            if (folderObject.DefineRole) {
                obj.CurrentApprover =
                    folderObject.ProjectmanagerEmail === null
                        ? folderObject.PublisherEmail
                        : folderObject.ProjectmanagerEmail;

                InternalStatus =
                    folderObject.ProjectmanagerEmail == null
                        ? "PendingWithPublisher"
                        : "PendingWithPM";
            }

            const status = await getStatusByInternalStatus(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                InternalStatus
            );

            obj.StatusId = status.value[0].ID;
            obj.InternalStatus = status.value[0].InternalStatus;
            obj.DisplayStatus = status.value[0].StatusName;

            const queryURL = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libName}')/items?$select=RefSequence,Created&$top=1&$orderby=RefSequence desc`;

            const LastDocRes = await getListData(queryURL, context);

            const lastSeq = LastDocRes.value[0]?.RefSequence ?? 0;

            const ReferenceNo = generateAutoRefNumber(
                lastSeq,
                folderObject,
                LastDocRes.value[0]?.Created,
                libName
            );

            obj.ReferenceNo = ReferenceNo.refNo.replace(/null/, "");
            obj.RefSequence = ReferenceNo.count;

            await updateLibrary(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                obj,
                itemId,
                libName
            );
            const fileInfoRes = await fetch(
                `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${fileServerRelativeUrl}')?$select=UniqueId`,
                {
                    headers: { Accept: "application/json;odata=verbose" }
                }
            );

            const fileInfo = await fileInfoRes.json();
            const uniqueId = fileInfo.d.UniqueId.replace(/-/g, "");

            const openUrl =
                `${context.pageContext.web.absoluteUrl}${fileServerRelativeUrl}` +
                `?d=w${uniqueId}`;

            window.open(openUrl, "_blank");

            setShowLoader({ display: "none" });
            dismissUploadPanel();


        } catch (error) {
            console.error("Error creating file:", error);

        }
    };


    const submit = async () => {
        let isValid = true;
        if (dynamicControl.length > 0) {
            dynamicControl.forEach((item: any) => {
                if (item.IsRequired && !dynamicValues[item.InternalTitleName]) {
                    setDynamicValuesErr((prev) => ({
                        ...prev,
                        [item.InternalTitleName]: DisplayLabel.ThisFieldisRequired,
                    }));
                    isValid = false;
                    return;
                }
            });
        }
        if (filetype === "upload" && attachmentsFiles.length === 0) {
            setAttachmentErr(DisplayLabel.ThisFieldisRequired);
            isValid = false;
        }
        if (!isValid) return;
        setShowLoader({ display: "block" });
        let count = 0;
        const queryURL = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libName}')/items?$select=EncodedAbsUrl,*,File/Name&$expand=File&$top=1&$orderby=RefSequence desc`;
        const LastDocRes = await getListData(queryURL, context);
        if (LastDocRes.value[0].RefSequence == null || LastDocRes.value[0].RefSequence == undefined) {
            LastDocRes.value[0].RefSequence = 0;
        }

        attachmentsFiles.forEach(async (item) => {
            if (item.isUpdateExistingFile === "Yes") {
                existingFile.map(async (el) => {
                    await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, { IsExistingFlag: "Old" }, el.ListItemAllFields.ID, libName);
                });
            }

            const Fileuniqueid = await uuidv4();
            let finalFileName = `${Fileuniqueid}-${item.attachment.name}`;
            const folderData = JSON.parse(JSON.stringify(folderObject, (key, value) => (value === null || (Array.isArray(value) && value.length === 0)) ? undefined : value));
            let obj: any = {
                ...folderData,
                ...dynamicValues,
                ActualName: item.attachment.name,
                FolderDocumentPath: `/${folderPath}`,
                OCRStatus: "Pending",
                UploadFlag: "Frontend",
                Level: item.version,
                IsExistingFlag: item.Flag,
                IsExistingRefID: item.IsExistingRefID,
            };
            let InternalStatus = "Published";

            if (folderObject.DefineRole) {
                obj.CurrentApprover = folderObject.ProjectmanagerEmail === null ? folderObject.PublisherEmail : folderObject.ProjectmanagerEmail;
                InternalStatus = folderObject.ProjectmanagerEmail == null ? "PendingWithPublisher" : "PendingWithPM";
            }

            const res = item.attachment.name.split('.').slice(0, -1).join('.');
            const extension = item.attachment.name.split('.').pop();
            const rename = (res).replace(/[^a-z0-9-\s]/gi, '');
            if (folderObject.DocumentSuffix !== null && folderObject.DocumentSuffix !== "") {
                let suffix = folderObject.DocumentSuffix;

                if (folderObject.DocumentSuffix === "Other") {
                    suffix = folderObject.OtherSuffix;
                }
                obj.ActualName = folderObject.PSType === "Prefix" ? `${suffix}_${rename}.${extension}` : obj.ActualName = `${rename}_${suffix}.${extension}`;
            }
            const status = await getStatusByInternalStatus(context.pageContext.web.absoluteUrl, context.spHttpClient, InternalStatus);

            obj.StatusId = status.value[0].ID;
            obj.InternalStatus = status.value[0].InternalStatus;
            obj.DisplayStatus = status.value[0].StatusName;
            obj.Active = true;
            const refCount = LastDocRes.value[0].RefSequence == null ? 0 : LastDocRes.value[0].RefSequence + count;
            const ReferenceNo = generateAutoRefNumber(refCount, folderObject, LastDocRes.value[0].Created, LibraryDetails);

            obj.ReferenceNo = ReferenceNo.refNo.replace(/null/, "");
            obj.RefSequence = ReferenceNo.count;

            if (item.Flag === "Replace") {

                const existingFile = FileData.find(
                    f => f.ListItemAllFields.ActualName?.toLowerCase() === item.attachment.name.toLowerCase()
                );

                if (existingFile) {
                    finalFileName = existingFile.Name;

                }
            }

            //  let UploadFileData = await UploadFile(context.pageContext.web.absoluteUrl, context.spHttpClient, item.attachment, `${Fileuniqueid}-${item.attachment.name}`, libName, obj, folderPath);
            let UploadFileData = await UploadFile(context.pageContext.web.absoluteUrl, context.spHttpClient, item.attachment, finalFileName, libName, obj, folderPath);

            console.log(UploadFileData);

            if (folderObject.DefineRole != null) {
                let emailObj: any = {
                    To: folderObject.ProjectmanagerEmail,
                    FolderPath: obj.FolderDocumentPath,
                    DocName: obj.ActualName,
                    AuthorTitle: context.pageContext.user.displayName,
                    TileName: libName,
                    Sub: DisplayLabel.PublisherEmailSubject + " " + obj.ReferenceNo,
                    Status: status.value[0].InternalStatus
                };
                emailObj.ID = folderObject.Id;
                emailObj.libraryName = libName;
                await TileSendMail(context, emailObj);
            }
            count++;

            if (item.IsExistingRefID !== "" && item.IsExistingRefID !== null && item.IsExistingRefID !== undefined) {
                if (LibraryDetails.IsArchiveRequired) {
                    const AllData = await getDataByRefID(context, item.IsExistingRefID, libName);
                    const ExistingRefData = AllData.value?.filter((ele: any) => ele.Active == true);
                    if (ExistingRefData?.length > archiveCount) {

                        const FileID = ExistingRefData[ExistingRefData?.length - 1].ID;
                        let updateArchiveObj = {
                            Active: false,
                            IsArchiveFlag: true
                        };

                        await updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, updateArchiveObj, FileID, libName);
                    }
                }
            }


            if (count === attachmentsFiles.length) {
                dismissUploadPanel();
                setShowLoader({ display: "none" });
                setAlertMsg(DisplayLabel.SubmitMsg);
                setIsPopupBoxVisible(true);
            }

        });
    };

    const hidePopup = useCallback(() => { setIsPopupBoxVisible(false); }, [isPopupBoxVisible]);

    return (
        <div>
            <Panel
                headerText={DisplayLabel.Upload}
                isOpen={isOpenUploadPanel}
                //onDismiss={dismissUploadPanel}
                onDismiss={(ev) => {
                    if (showConfirmDialog) return;
                    dismissUploadPanel();
                }}
                isLightDismiss={false}
                isBlocking={true}
                closeButtonAriaLabel="Close"
                type={PanelType.large}
                onRenderFooterContent={() => (<>
                    {filetype === "upload" ? (
                        <PrimaryButton onClick={submit} styles={{ root: { marginRight: 8 } }} disabled={showLoader.display === "block"}>{DisplayLabel.Submit}</PrimaryButton>
                    ) : (
                        <PrimaryButton onClick={createOfficeFile} styles={{ root: { marginRight: 8 } }} disabled={showLoader.display === "block"}>Create File</PrimaryButton>
                    )}
                    <DefaultButton onClick={dismissUploadPanel} disabled={showLoader.display === "block"}>{DisplayLabel.Cancel}</DefaultButton>
                </>)}
                isFooterAtBottom={true}
            >
                <div>
                    <div className="row">
                        <div className="column12">
                            <label>{DisplayLabel.Path}: <b>{folderPath}</b></label>
                        </div>
                    </div>
                    <div className="row">
                        <div className="column6">
                            <TextField label={DisplayLabel.TileName} value={libName} readOnly />
                        </div>
                        <div className="column6">
                            <TextField label={DisplayLabel.FolderName} value={folderName} readOnly />
                        </div>
                    </div>
                    <div className="row">
                        {renderDynamicControls()}
                    </div>
                    {filetype === "upload" && (
                        <div className="grid upload-file-row">
                            <div className="grid-item-large">
                                <Field className="fluent-file-upload">
                                    <div className="fluent-file-upload-control">
                                        <label>{DisplayLabel.ChooseFile}<span style={{ color: "red" }}>*</span> </label>
                                        <label style={{ color: "red" }}>
                                            {DisplayLabel?.FileAttachmentNote || ""}
                                        </label>
                                        <input
                                            type="file"
                                            multiple
                                            onChange={(event: React.ChangeEvent<HTMLInputElement>) => {
                                                if (event.target.files) {
                                                    setAttachment(Array.from(event.target.files));
                                                }
                                            }}
                                            key={fileKey}
                                        />
                                    </div>
                                    <span style={{ color: "red" }}>{attachmentErr}</span>
                                </Field>

                            </div>
                            <div className="grid-item-small upload-file-action">
                                <IconButton
                                    iconProps={{ iconName: 'Add' }}
                                    style={{ background: "#009ef7", color: "#fff", border: "#009ef7" }}
                                    onClick={addAttachment}
                                    label="Add"
                                />
                            </div>
                        </div>)}
                    <div className="row">
                        {attachmentsFiles.length ?
                            <table className="fluent-table">
                                <thead>
                                    <tr>
                                        <th style={{ width: 100 }}>{DisplayLabel.SrNo}</th>
                                        <th>{DisplayLabel.FileName}</th>
                                        <th>{DisplayLabel.IsthisAnUpdateToExistingFile}</th>
                                        {isUpdateExistingFile && <th>{DisplayLabel.FileName}</th>}
                                        <th>{DisplayLabel.Versions}</th>
                                        <th style={{ width: 100 }}>{DisplayLabel.Action}</th>
                                    </tr>
                                </thead>

                                <tbody>
                                    {attachmentsFiles?.map((item, index) => (
                                        <tr key={index}>
                                            <td>{index + 1}</td>

                                            <td>{item.attachment.name}</td>

                                            <td style={{ width: 200 }}>


                                                <Select options={[
                                                    { value: "Yes", label: 'Yes' },
                                                    { value: "No", label: 'No' },
                                                ]}

                                                    value={{ value: item.isUpdateExistingFile, label: item.isUpdateExistingFile }}
                                                    isDisabled={item.isDisabled}
                                                    onChange={async (option: any) => {
                                                        const attach = await Promise.all(attachmentsFiles.map((ele, i) => i === index ? { ...ele, isUpdateExistingFile: option?.value } : ele));
                                                        let filterFiles = files.filter((el: any) => el.ListItemAllFields.IsExistingFlag === "New");
                                                        if (filterFiles.length > 0 && attachmentsFiles.length > 0) {
                                                            attachmentsFiles.map((el: any) => {
                                                                filterFiles = filterFiles.filter((ele: any) => {
                                                                    if (el.name !== "" && item.name != el.name) {
                                                                        return ele.ListItemAllFields.Active === true && ele.ListItemAllFields.IsExistingFlag === "New" && el.name != ele.Name;
                                                                    } else {
                                                                        return ele.ListItemAllFields.Active === true && ele.ListItemAllFields.IsExistingFlag === "New";
                                                                    }
                                                                });
                                                            });
                                                        }
                                                        setFilterFilesData(filterFiles);
                                                        setFilesData(filterFiles.map((el: any) => ({ value: el.Name, label: el.ListItemAllFields.ActualName })));
                                                        setAttachmentsFiles(attach);
                                                        const filterD = attach.filter((el, i) => el.isUpdateExistingFile === "Yes");
                                                        filterD.length > 0 ? setIsUpdateExistingFile(option?.value === "Yes") : "";
                                                    }} />
                                            </td>

                                            {isUpdateExistingFile && (
                                                <td style={{ width: 250 }}>
                                                    <Select
                                                        options={filesData}
                                                        value={filesData.find(
                                                            (option: any) => option.value === item.OldFileName
                                                        )}
                                                        isSearchable
                                                        placeholder={DisplayLabel?.Selectanoption}
                                                        isDisabled={item.isDisabled}
                                                        onChange={(option: any) => {
                                                            const fData = filterFilesData.filter(
                                                                (ele: any) => ele.Name === option?.value
                                                            );

                                                            let level = 1.0;

                                                            if (fData.length > 0)
                                                                level = parseFloat(fData[0].ListItemAllFields.Level) + 1.0;

                                                            setAttachmentsFiles((prev) =>
                                                                prev.map((ele, i) =>
                                                                    i === index
                                                                        ? {
                                                                            ...ele,
                                                                            OldFileName: option?.value,
                                                                            version: level.toFixed(1)
                                                                        }
                                                                        : ele
                                                                )
                                                            );
                                                        }}
                                                    />
                                                </td>
                                            )}

                                            <td style={{ width: 120 }}>
                                                <TextField value={item.version} disabled />
                                            </td>

                                            <td>
                                                <IconButton
                                                    iconProps={{ iconName: item.isDisabled ? "Edit" : "Save" }}
                                                    styles={{ root: { color: "#009ef7" } }}
                                                    onClick={() => onClickDetails(index)}
                                                />

                                                <IconButton
                                                    iconProps={{ iconName: "Delete" }}
                                                    styles={{ root: { color: "red" } }}
                                                    onClick={() =>
                                                        setAttachmentsFiles((prev) =>
                                                            prev.filter((ele, i) => i !== index)
                                                        )
                                                    }
                                                />
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                            : <></>
                        }
                    </div>


                    {filetype !== "upload" && (
                        <div className="row">
                            <div className="column12">
                                <TextField
                                    label="File Name"
                                    value={newFileName}
                                    onChange={(e, val) => setNewFileName(val || "")}
                                    required
                                    errorMessage={fileNameError}
                                />
                            </div>
                        </div>
                    )}
                </div>
                {showLoader.display === "block" && (
                    <div
                        style={{
                            position: "fixed",
                            inset: 0,
                            background: "rgba(255, 255, 255, 0.72)",
                            backdropFilter: "blur(2px)",
                            zIndex: 100000,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center"
                        }}
                    >
                        <PageLoader message="Uploading file..." minHeight="auto" />
                    </div>
                )}
            </Panel>
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} type="insert" />
            <ConfirmationDialog
                hideDialog={showConfirmDialog}
                closeDialog={() => setShowConfirmDialog(false)}
                handleConfirm={handleDuplicateConfirm}
                msg={`A file with this name already exists so we couldn't upload ${existingFileNamesInFileData.join(", ")}.Add it as a new version of the existing file, or keep them both.`}
                Yes="Keep Both"
                No="Replace"
            />


        </div>
    );
}

export default React.memo(UploadFiles);
