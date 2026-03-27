import React, { memo, useCallback, useEffect, useRef, useState } from "react";
import {
    ChoiceGroup,
    DefaultButton,
    Panel,
    PanelType,
    PrimaryButton,
    TextField,
    Toggle,
    DatePicker, mergeStyleSets
} from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
    IPeoplePickerContext,
    PeoplePicker,
    PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Field, Label } from "@fluentui/react-components";
import { getActiveTypeData } from "../../../../Services/PrefixSuffixMasterService";
import { getConfigActive } from "../../../../Services/ConfigService";
import { getDataByLibraryName } from "../../../../Services/MasTileService";
import { getAllFolder, getListData, updateLibrary } from "../../../../Services/GeneralDocument";
import { FolderStructure } from "../../../../Services/FolderStructure";
import { getUserIdFromLoginName } from "../../../../DAL/Commonfile";
import PopupBox from "./PopupBox";
import PageLoader from "./PageLoader";
import FieldError from "./FieldError";
import { ILabel } from "../../../../Intrface/ILabel";
import Select from 'react-select';
import { getTemplateActive } from "../../../../Services/TemplateService";
import { getActiveFolder } from "../../../../Services/FolderMasterService";
import { format } from "date-fns";
//import { SPHttpClient } from "@microsoft/sp-http";


export interface IProjectEntryProps {
    isOpen: boolean;
    dismissPanel: (value: boolean) => void;
    context: WebPartContext;
    LibraryDetails: any;
    admin: any;
    FormType: string;
    folderObject: any;
    folderPath: string;
    ChildFolderRoleInheritance: boolean;
}


const ProjectEntryForm: React.FC<IProjectEntryProps> = ({
    isOpen,
    dismissPanel,
    context,
    LibraryDetails,
    admin,
    FormType,
    folderObject,
    folderPath,
    ChildFolderRoleInheritance
}) => {
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const [folderName, setFolderName] = useState<string>("");
    const [isSuffixRequired, setIsSuffixRequired] = useState<boolean>(false);
    const [SuffixData, setSuffixData] = useState<any[]>([]);
    const [Suffix, setSuffix] = useState<string>("");
    const [OtherSuffix, setOtherSuffix] = useState<string>("");
    const [configData, setConfigData] = useState<any[]>([]);
    const [dynamicControl, setDynamicControl] = useState<any[]>([]);
    const [libraryDetails, setLibraryDetails] = useState<any>({});
    const [options, setOptions] = useState<any>({});
    const [dynamicValues, setDynamicValues] = useState<{ [key: string]: any; }>({});
    const buttonStyles = { root: { marginRight: 8 } };
    const [folderAccess, setFolderAccess] = useState<any[]>([]);
    const [usersIds, setUsersIds] = useState<any[]>([]);
    const [publisher, setPublisher] = useState<any[]>([]);
    const [approver, setApprover] = useState<any[]>([]);
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState(false);
    const [popupType, setPopupType] = useState<"success" | "warning" | "insert" | "checkin" | "checkout" | "approve" | "reject" | "delete" | "update" | "restore" | "grant" | "remove">("success");
    const [alertMsg, setAlertMsg] = useState("");
    const [isApprovalRequired, setIsApprovalRequired] = useState<boolean>(false);
    const [allUsers, setAllUsers] = useState<any>([]);

    const [folderNameErr, setFolderNameErr] = useState<string>("");
    const [SuffixErr, setSuffixErr] = useState<string>("");
    const [OtherSuffixErr, setOtherSuffixErr] = useState<string>("");
    const [dynamicValuesErr, setDynamicValuesErr] = useState<{ [key: string]: string; }>({});
    const [folderAccessErr, setFolderAccessErr] = useState<string>("");
    const [publisherErr, setPublisherErr] = useState<string>("");
    const [approverErr, setApproverErr] = useState<string>("");
    const [showLoader, setShowLoader] = useState(false);
    const [isDisabled, setIsDisabled] = useState<boolean>(false);
    const [projectManagerEmail, setProjectManagerEmail] = useState("");
    const [publisherEmail, setPublisherEmail] = useState("");
    const [panelTitle, setPanelTitle] = useState(DisplayLabel.EntryForm);
    const [createStructure, setCreateStructure] = useState<boolean>(false);
    const [allFolderTemplate, setAllFolderTemplate] = useState<any>([]);
    const [folderTemplate, setFolderTemplate] = useState<any>("");
    const [folderTemplateErr, setFolderTemplateErr] = useState<any>("");
    const [folderStructure, setFolderStructure] = useState<any>([]);
    const inputRefs = useRef<{ [key: string]: HTMLInputElement | null; }>({});
    const [TemFolderName, setTemFolderName] = useState<string>("");
    const [isInitialLoading, setIsInitialLoading] = useState(true);

    const meargestyles = mergeStyleSets({
        root: { selectors: { '> *': { marginBottom: 15 } } },
        control: { maxWidth: "100%", marginBottom: 15 },
    });

    const handleInputChange = (fieldName: string, value: any) => {
        setDynamicValues((prevValues) => ({
            ...prevValues,
            [fieldName]: value,
        }));
    };

    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any
    };


    const handleToggleChange = (_: React.MouseEvent<HTMLElement>, checked?: boolean) => {
        setIsSuffixRequired(!!checked);
    };

    useEffect(() => {
        Promise.all([
            fetchLibraryDetails(),
            fetchSuffixData(),
            getAllUsers(),
            getFolderStructure(),
            getFolderTemplate()
        ]).finally(() => setIsInitialLoading(false));
    }, []);

    useEffect(() => {
        setCreateStructure(false);
        setFolderTemplate("");
        clearErr();
        clearFeilds();
        setIsDisabled(FormType === "ViewForm");
        FormType !== "EntryForm" ? bindFormData() : "";
        if (FormType === "ViewForm")
            setPanelTitle(DisplayLabel.ViewForm);
        else if (FormType === "EditForm")
            setPanelTitle(DisplayLabel.EditForm);
        else
            setPanelTitle(DisplayLabel.EntryForm);

    }, [isOpen]);

    const getFolderTemplate = async () => {
        const data = await getTemplateActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
        if (data.value.length > 0)
            setAllFolderTemplate(data.value.map((el: any) => ({ value: el.Name, label: el.Name })));
    };

    const getFolderStructure = async () => {
        const data = await getActiveFolder(context.pageContext.web.absoluteUrl, context.spHttpClient);
        setFolderStructure(data.value);
    };

    const getAllUsers = async () => {
        const data = await getListData(`${context.pageContext.web.absoluteUrl}/_api/web/siteusers?$filter=PrincipalType eq 1`, context);
        if (data.value.length > 0) {
            setAllUsers(data.value);
        }
    };

    const fetchSuffixData = async () => {
        const data = await getActiveTypeData(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            "Suffix"
        );
        const column = data.value.map((item: any) => ({
            value: item.PSName,
            label: item.PSName,
        }));
        setSuffixData(column);
    };


    const fetchLibraryDetails = async () => {
        const dataConfig = await getConfigActive(context.pageContext.web.absoluteUrl, context.spHttpClient);
        const libraryData = await getDataByLibraryName(context.pageContext.web.absoluteUrl, context.spHttpClient, LibraryDetails?.LibraryName);

        setLibraryDetails(libraryData.value[0]);
        setConfigData(dataConfig.value);

        if (libraryData.value[0]?.DynamicControl) {
            let jsonData = JSON.parse(libraryData.value[0].DynamicControl);
            jsonData = jsonData.filter((ele: any) => ele.IsActiveControl);
            jsonData = jsonData.map((el: any) => {
                if (el.ColumnType === "Person or Group") {
                    el.InternalTitleName = `${el.InternalTitleName}Id`;
                }
                return el;
            });
            setDynamicControl(jsonData);
            bindDropdown(jsonData);
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
                    const data = await getListData(
                        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${item.InternalListName}')/items?$top=5000&$filter=Active eq 1&$orderby=${item.DisplayValue} asc`,
                        context
                    );
                    dropdownOptions = data.value.map((ele: any) => ({
                        value: ele[item.DisplayValue],
                        label: ele[item.DisplayValue],
                    }));
                }
                setOptions((prev: any) => ({ ...prev, [item.InternalTitleName]: dropdownOptions }));
            }
        });
    };

    const renderDynamicControls = useCallback(() => {
        return dynamicControl.filter((item: any, index: number) => !item.IsFieldAllowInFile).map((item: any, index: number) => {
            const filterObj = configData.find((ele) => ele.Id === item.Id);

            if (!filterObj) return null;

            switch (item.ColumnType) {
                case "Dropdown":
                case "Multiple Select":
                    return (
                        <Field key={index}>
                            <Label required={item.IsRequired} >{item.Title}</Label>
                            <Select
                                options={options[item.InternalTitleName]}
                                required={item.IsRequired}
                                value={(options[item.InternalTitleName] || []).find((option: any) => option.value === dynamicValues[item.InternalTitleName])}
                                onChange={(option: any) => handleInputChange(item.InternalTitleName, option?.value)}
                                isSearchable
                                placeholder={DisplayLabel?.Selectanoption}
                                isMulti={item.ColumnType === "Multiple Select"}
                                isDisabled={isDisabled}
                                ref={(input: any) => (inputRefs.current[item.InternalTitleName] = input)}
                            />
                            <FieldError message={dynamicValuesErr[item.InternalTitleName]} />
                        </Field>
                    );

                case "Person or Group":
                    return (
                        // <div className={dynamicControl.length > 5 ? styles.col6 : styles.col12} key={index}>
                        <Field key={index}>
                            <PeoplePicker
                                titleText={item.Title}
                                context={peoplePickerContext}
                                personSelectionLimit={20}
                                showtooltip={true}
                                required={item.IsRequired}
                                showHiddenInUI={false}
                                principalTypes={[PrincipalType.User]}
                                // onChange={(items) => handleInputChange(item.InternalTitleName, items)}
                                onChange={async (items) => {
                                    try {
                                        const userIds = await Promise.all(
                                            items.map(async (item: any) => {
                                                const data = await getUserIdFromLoginName(context, item.id);
                                                return data.Id;
                                            })
                                        );
                                        setDynamicValues((prevValues) => ({
                                            ...prevValues,
                                            [item.InternalTitleName]: userIds[0],
                                        }));
                                        setUsersIds((prev) => [...prev, ...userIds]);
                                    } catch (error) {
                                        console.error("Error fetching user IDs:", error);
                                    }
                                }}
                                defaultSelectedUsers={allUsers
                                    .filter((el: any) => el.Id === dynamicValues[item.InternalTitleName])
                                    .map((user: any) => user.Email)}
                                disabled={isDisabled}
                                ref={(input: any) => (inputRefs.current[item.InternalTitleName] = input)}
                            />
                            <FieldError message={dynamicValuesErr[item.InternalTitleName]} />
                        </Field>
                    );

                case "Radio":
                    const radioOptions = filterObj.StaticDataObject.split(";").map((ele: string) => ({
                        key: ele,
                        text: ele,
                    }));
                    return (
                        // <div className={dynamicControl.length > 5 ? styles.col6 : styles.col12} key={index}>
                        <Field key={index}>
                            <ChoiceGroup
                                options={radioOptions}
                                onChange={(ev, option) => handleInputChange(item.InternalTitleName, option?.key)}
                                selectedKey={dynamicValues[item.InternalTitleName] || ""}
                                label={item.Title}
                                required={item.IsRequired}
                                disabled={isDisabled}
                            />
                        </Field>
                    );
                case "Date and Time":
                    return (
                        //<div className={dynamicControl.length > 5 ? styles.col6 : styles.col12} key={index}>
                        <Field key={index}>
                            <label >{item.Title}{item.IsRequired ? <span style={{ color: "red" }}>*</span> : <></>}</label>
                            <DatePicker
                                componentRef={(input: any) => (inputRefs.current[item.InternalTitleName] = input)}
                                onSelectDate={(date: Date | null | undefined) => handleInputChange(item.InternalTitleName, date)}
                                className={meargestyles.control}
                                value={dynamicValues[item.InternalTitleName] || ""}
                                disabled={isDisabled}
                                formatDate={(date) => date ? format(new Date(date), "dd/MM/yyyy") : ''}
                            />
                            <FieldError message={dynamicValuesErr[item.InternalTitleName]} />
                        </Field>
                    );

                default:
                    return (
                        //<div className={dynamicControl.length > 5 ? styles.col6 : styles.col12} key={index}>
                        <Field key={index}>
                            <TextField
                                type={"text"}
                                label={item.Title}
                                value={dynamicValues[item.InternalTitleName] || ""}
                                onChange={(ev, value) => handleInputChange(item.InternalTitleName, removeSepcialCharacters(value))}
                                multiline={item.ColumnType === "Multiple lines of Text"}
                                required={item.IsRequired}
                                errorMessage={dynamicValuesErr[item.InternalTitleName]}
                                disabled={isDisabled}
                                componentRef={(input: any) => (inputRefs.current[item.InternalTitleName] = input)}
                            />
                        </Field>
                    );
            }
        });
    }, [dynamicControl, options, dynamicValues, dynamicValuesErr]);

    const clearErr = () => {
        setFolderNameErr("");
        setApproverErr("");
        setPublisherErr("");
        setFolderAccessErr("");
        setDynamicValuesErr({});
        setSuffixErr("");
        setOtherSuffixErr("");
        setFolderTemplateErr("");
    };

    const clearFeilds = () => {
        setFolderName("");
        setSuffix("");
        setOtherSuffix("");
        setDynamicValues({});
        setFolderAccess([]);
        setPublisher([]);
        setApprover([]);
        setIsApprovalRequired(false);
        setIsSuffixRequired(false);
    };

    const submit = async () => {
        // e.preventDefault();
        clearErr();
        let isValid = true;
        if (folderName.trim() === "") {
            setFolderNameErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["FolderName"]?.focus();
            isValid = false;
            return;
        }
        if (isSuffixRequired && Suffix === "") {
            setSuffixErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["Suffix"]?.focus();
            isValid = false;
            return;
        }
        if (Suffix === "Other" && OtherSuffix.trim() === "") {
            setOtherSuffixErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["OtherSuffix"]?.focus();
            isValid = false;
            return;
        }

        if (dynamicControl.length > 0) {
            dynamicControl.filter((item: any, index: number) => !item.IsFieldAllowInFile).forEach((item: any) => {
                if (item.IsRequired && !dynamicValues[item.InternalTitleName]) {
                    setDynamicValuesErr((prev) => ({
                        ...prev,
                        [item.InternalTitleName]: DisplayLabel.ThisFieldisRequired,
                    }));
                    inputRefs.current[item.InternalTitleName]?.focus();
                    isValid = false;
                    return;
                }
            });
        }

        if (FormType === "EntryForm" && folderAccess.length === 0) {
            setFolderAccessErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["FolderAccess"]?.focus();
            return;
        }
        if (isApprovalRequired && approver.length === 0) {
            setApproverErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["Approver"]?.focus();
            isValid = false;
            return;
        }
        if (isApprovalRequired && publisher.length === 0) {
            setPublisherErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["Publisher"]?.focus();
            isValid = false;
            return;
        }
        if (createStructure && folderTemplate === "") {
            setFolderTemplateErr(DisplayLabel.ThisFieldisRequired);
            inputRefs.current["FolderTemplate"]?.focus();
            isValid = false;
            return;
        }
        if (FormType === "EntryForm") {
            const data = await getAllFolder(context.pageContext.web.absoluteUrl, context, LibraryDetails.LibraryName);
            if (data && data.Folders.filter((el: any) => el.Name === folderName).length > 0) {
                setFolderNameErr(DisplayLabel.FolderAlreadyExist);
                inputRefs.current["FolderName"]?.focus();
                isValid = false;
                return;
            }
        }
        if (isValid)
            createFolder();
    };



    const createFolder = async () => {
        setShowLoader(true);
        if (FormType === "EntryForm") {

            const users = [
                ...folderAccess.map(id => ({ id, type: 'FolderAccess' })),
                ...usersIds.map(id => ({ id, type: 'User' })),
                ...publisher.map(id => ({ id, type: 'Publisher' })),
                ...approver.map(id => ({ id, type: 'Approver' })),
                ...admin.map((id: any) => ({ id, type: 'Admin' })),
                ...(LibraryDetails.TileAdminId
                    ? [{ id: LibraryDetails.TileAdminId, type: 'TileAdmin' }]
                    : []),
            ];



            console.log(users);

            FolderStructure(context, `${LibraryDetails.LibraryName}/${folderName}`, users, LibraryDetails.LibraryName, true).then(async (response) => {
                console.log(response);
                await updateFolderMetaData(response);
                if (createStructure) {
                    if (TemFolderName === "") {
                        createFolderStructure(users);
                    }
                    else {
                        FolderStructure(context, `${LibraryDetails.LibraryName}/${folderName}/${TemFolderName}`, users, LibraryDetails.LibraryName, ChildFolderRoleInheritance).then(async (response) => {
                            await updateFolderMetaData(response);
                            createFolderStructure(users);
                        });
                    }
                }
            });
        }
        else {
            await updateFolderMetaData(folderObject.id);
            const folders = await getAllFolder(context.pageContext.web.absoluteUrl, context, folderPath);
            folders.Folders.map((folder: any) => { updateFolderMetaData(folder.ListItemAllFields.Id); });

        }
    };
    const updateFolderMetaData = async (id: number) => {
        let obj: any = {
            ...dynamicValues,
            DocumentSuffix: Suffix || "",
            OtherSuffix: OtherSuffix || "",
            IsSuffixRequired: isSuffixRequired,
            PSType: "Suffix",
            DefineRole: isApprovalRequired,
            CreateFolder: createStructure,
            Template: folderTemplate,
        };
        if (isApprovalRequired) {
            const filterApprover = allUsers.filter((el: any) => el.Id === approver[0])[0];
            obj.ProjectmanagerAllow = true;
            obj.ProjectmanagerId = filterApprover.Id;
            obj.ProjectmanagerEmail = filterApprover.Email;
            const filterPublisher = allUsers.filter((el: any) => el.Id === publisher[0])[0];
            obj.PublisherAllow = true;
            obj.PublisherId = filterPublisher.Id;
            obj.PublisherEmail = filterPublisher.Email;
        }

        updateLibrary(context.pageContext.web.absoluteUrl, context.spHttpClient, obj, id, LibraryDetails.LibraryName).then((response) => {
            if (!createStructure) {
                dismissPanel(false);
                setShowLoader(false);
                setAlertMsg(DisplayLabel.FolderUpdatedMsg);
                setPopupType("update");
                setIsPopupBoxVisible(true);
            }
        });
    };

    const createFolderStructure = async (users: any) => {
        //  const filterFolders = folderStructure.filter((el: any) => el.TemplateName.Name === folderTemplate);
        const filterFolders = folderStructure.filter(
            (el: any) => el.TemplateName?.Name === folderTemplate
        );
        const firstlevel = getFirstLevel(filterFolders);
        let count = 0;
        const Updatedfolderpath = TemFolderName ? `${folderName}/${TemFolderName}` : folderName;
        firstlevel.map(async (folder: any) => {
            const response = await FolderStructure(context, `${LibraryDetails.LibraryName}/${Updatedfolderpath}/${folder.FolderName}`, users, LibraryDetails.LibraryName, true);
            await updateFolderMetaData(response);
            const ChildLevel = getEqualToData(filterFolders, folder.Id);
            await createChildFolder(ChildLevel, folder.FolderName, users);
            count++;
            if (firstlevel.length === count) {
                dismissPanel(false);
                setShowLoader(false);
                setAlertMsg(DisplayLabel.SubmitMsg);
                setPopupType("insert");
                setIsPopupBoxVisible(true);
            }
        });
    };

    const createChildFolder = async (folder: any, Name: any, users: any) => {
        const basePath = TemFolderName ? `${folderName}/${TemFolderName}` : folderName;
        folder.map(async (folder: any) => {
            const ChildLevel = getEqualToData(folderStructure, folder.Id);
            if (ChildLevel.length > 0) {
                const response = await FolderStructure(context, `${LibraryDetails.LibraryName}/${basePath}/${Name}/${folder.FolderName}`, users, LibraryDetails.LibraryName, ChildFolderRoleInheritance);
                await updateFolderMetaData(response);
                await createChildFolder(ChildLevel, `${Name}/${folder.FolderName}`, users);
            }
            else {
                const response = await FolderStructure(context, `${LibraryDetails.LibraryName}/${basePath}/${Name}/${folder.FolderName}`, users, LibraryDetails.LibraryName, ChildFolderRoleInheritance);
                await updateFolderMetaData(response);
            }
        });
    };

    function getFirstLevel(item: any) {
        return item.filter((it: any) => it.ParentFolderIdId == null);
    }

    function getEqualToData(Folders: any, id: number) {
        return Folders.filter((it: any) => it.ParentFolderIdId === id);
    }

    const bindFormData = () => {
        setFolderName(folderObject.name);
        if (folderObject?.CreateFolder === true) {
            setTemFolderName(folderObject?.TemFolderName);
        }
        setIsSuffixRequired(folderObject?.IsSuffixRequired);
        if (folderObject?.IsSuffixRequired) {
            setSuffix(folderObject?.DocumentSuffix);
            folderObject?.DocumentSuffix === "Other" ? setOtherSuffix(folderObject?.OtherSuffix) : "";
        }
        setCreateStructure(folderObject?.CreateFolder);
        setFolderTemplate(folderObject?.Template);
        if (libraryDetails.AllowApprover) {
            setIsApprovalRequired(folderObject?.DefineRole);
            if (folderObject?.DefineRole) {
                setProjectManagerEmail(folderObject?.ProjectmanagerEmail);
                setPublisherEmail(folderObject?.PublisherEmail);
                setApprover([folderObject?.ProjectmanagerId]);
                setPublisher([folderObject?.PublisherId]);
            }
        }

        dynamicControl.map((item: any, index: number) => {
            const filterObj = configData.find((ele) => ele.Id === item.Id);
            if (!filterObj) return null;

            setDynamicValues((prevValues) => {
                let value = folderObject[item.InternalTitleName];
                filterObj.ColumnType === "Date and Time" ? value = new Date(value) : "";
                return {
                    ...prevValues,
                    [item.InternalTitleName]: value,
                };
            });
        });

    };

    const hidePopup = useCallback(() => { setIsPopupBoxVisible(false); }, [isPopupBoxVisible]);
    const isValidNumberString = (value: string): boolean => {
        return !isNaN(Number(value)) && value.trim() !== "";
    };
    const removeSepcialCharacters = (newValue?: string) => newValue?.replace(/[^a-zA-Z0-9\s]/g, '') || '';
    const removeFolderSepcialCharacters = (newValue?: string) =>
        newValue?.replace(/[^a-zA-Z0-9_\-\s]/g, '') || '';

    return (
        <>
            <Panel
                headerText={panelTitle}
                isOpen={isOpen}
                onDismiss={() => dismissPanel(false)}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                onRenderFooterContent={() => (
                    <>
                        {FormType !== "ViewForm" ? <PrimaryButton onClick={submit} styles={buttonStyles} >{FormType === "EntryForm" ? DisplayLabel.Submit : DisplayLabel.Update}</PrimaryButton> : <></>}
                        <DefaultButton onClick={() => dismissPanel(false)}>{DisplayLabel.Cancel}</DefaultButton>
                    </>
                )}
                isFooterAtBottom={true}
            >
                <div style={{ position: "relative" }}>
                    {isInitialLoading ? <PageLoader message="Loading form..." minHeight="52vh" /> : <div className="grid-2">
                        <div className="col-md-6">
                            <TextField
                                label={DisplayLabel.TileName}
                                value={LibraryDetails?.TileName}
                                disabled
                            />
                        </div>
                        <div className="col-md-6">
                            <TextField
                                value={folderName}
                                label={DisplayLabel.FolderName}
                                required
                                onChange={(_, newValue) => {
                                    const validName = removeFolderSepcialCharacters(newValue);
                                    setFolderName(validName);
                                }}
                                disabled={isDisabled || FormType === "EditForm"}
                            />
                            <FieldError message={folderNameErr} />

                        </div>
                    </div>}

                    <div className="grid-2">
                        <div className="col-12">
                            <Toggle
                                label={DisplayLabel.IsSuffixRequired}
                                onChange={handleToggleChange}
                                checked={isSuffixRequired}
                                disabled={isDisabled}
                            />
                        </div>
                    </div>

                    {isSuffixRequired && (
                        <>

                            <Field label={DisplayLabel.DocumentSuffix} required>
                                <Select
                                    options={SuffixData}
                                    value={SuffixData.find((option: any) => option.value === Suffix)}
                                    onChange={(option: any) => setSuffix(option.value as string)}
                                    isSearchable
                                    placeholder={DisplayLabel?.Selectanoption}
                                    isDisabled={isDisabled}
                                    ref={(input: any) => (inputRefs.current["Suffix"] = input)}
                                />
                                <FieldError message={SuffixErr} />
                            </Field>

                            {Suffix === "Other" && (
                                <Field>
                                    <TextField
                                        label={DisplayLabel.OtherSuffixName}
                                        value={OtherSuffix}
                                        onChange={(_, newValue) =>
                                            setOtherSuffix(removeSepcialCharacters(newValue))
                                        }
                                        required
                                        disabled={isDisabled}
                                    />
                                    <span className="errorText">{OtherSuffixErr}</span>
                                </Field>

                            )}
                        </>
                    )}

                    <div >{renderDynamicControls()}</div>
                    {libraryDetails?.AllowApprover ? <>
                        <Field>
                            <Toggle
                                label={DisplayLabel.IsApprovalFlowRequired}
                                onChange={() => { setIsApprovalRequired((pre) => !pre); }}
                                disabled={isDisabled}
                                checked={isApprovalRequired}
                            />
                        </Field>
                        {
                            isApprovalRequired ?
                                <div className="grid-2">
                                    <Field>
                                        <PeoplePicker
                                            titleText={DisplayLabel.Approver}
                                            context={peoplePickerContext}
                                            personSelectionLimit={1}
                                            showtooltip={true}
                                            required
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            ensureUser={true}
                                            onChange={async (items: any) => {
                                                try {
                                                    setProjectManagerEmail(items[0].secondaryText as string);
                                                    setApprover([items[0].id]);
                                                } catch (error) {
                                                    console.error("Error fetching user IDs:", error);
                                                }
                                            }}
                                            defaultSelectedUsers={[projectManagerEmail]}
                                            disabled={isDisabled}
                                            ref={(input: any) => (inputRefs.current["Approver"] = input)}
                                        />
                                        <FieldError message={approverErr} />
                                    </Field>
                                    <Field >
                                        <PeoplePicker
                                            titleText={DisplayLabel.Publisher}
                                            context={peoplePickerContext}
                                            personSelectionLimit={1}
                                            showtooltip={true}
                                            required
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User]}
                                            defaultSelectedUsers={[publisherEmail]}
                                            ensureUser={true}
                                            onChange={async (items: any) => {
                                                try {
                                                    setPublisherEmail(items[0].secondaryText as string);
                                                    setPublisher([items[0].id]);
                                                } catch (error) {
                                                    console.error("Error fetching user IDs:", error);
                                                }
                                            }}
                                            disabled={isDisabled}
                                            ref={(input: any) => (inputRefs.current["Publisher"] = input)}
                                        />
                                        <FieldError message={publisherErr} />

                                    </Field>
                                </div> : <></>
                        }
                    </> : <></>}
                    <Field>
                        {FormType === "EntryForm" ?
                            <>
                                <PeoplePicker
                                    titleText={DisplayLabel.FolderAccess}
                                    context={peoplePickerContext}
                                    personSelectionLimit={20}
                                    showtooltip={true}
                                    required
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup]}
                                    onChange={async (items: any[]) => {
                                        try {
                                            const userIds = await Promise.all(
                                                items.map(async (item: any) => {
                                                    let userid: number = 0;
                                                    if (isValidNumberString(item.id)) {
                                                        userid = Number(item.id);
                                                    } else {
                                                        const data = await getUserIdFromLoginName(context, item.id);
                                                        userid = data.Id;
                                                    };
                                                    return userid;
                                                })
                                            );
                                            setFolderAccess(userIds);
                                        } catch (error) {
                                            console.error("Error fetching user IDs:", error);
                                        }
                                    }}
                                    ref={(input: any) => (inputRefs.current["FolderAccess"] = input)}
                                />
                                <FieldError message={folderAccessErr} />
                            </>
                            : <></>
                        }
                    </Field>
                    <div className="grid-2">
                        <Field >
                            <Toggle
                                label={DisplayLabel.CreateStructure}
                                onChange={() => { setCreateStructure((pre) => !pre); }}
                                disabled={isDisabled || FormType === "EditForm"}
                                checked={createStructure}
                            />
                        </Field>
                        {
                            createStructure ?
                                <Field>
                                    <TextField
                                        label={DisplayLabel.FolderName}
                                        value={TemFolderName}
                                        onChange={(ev, newValue) => setTemFolderName(newValue || "")}
                                        disabled={isDisabled || FormType === "EditForm"}
                                    />
                                </Field> : <></>
                        }
                    </div>

                    <div className="grid-2">
                        {
                            createStructure ?
                                <Field label={DisplayLabel.TemplateName} required>
                                    <Select
                                        options={allFolderTemplate}
                                        value={allFolderTemplate.find((option: any) => option.value === folderTemplate)}
                                        onChange={(option: any) => setFolderTemplate(option.value as string)}
                                        isSearchable
                                        placeholder={DisplayLabel?.Selectanoption}
                                        ref={(input: any) => (inputRefs.current["CreateStructure"] = input)}
                                        isDisabled={isDisabled || FormType === "EditForm"}
                                    />
                                    <FieldError message={folderTemplateErr} />
                                </Field> : <></>
                        }
                    </div>

                    {showLoader && (
                        <div
                            style={{
                                position: "absolute",
                                inset: 0,
                                background: "rgba(255, 255, 255, 0.78)",
                                backdropFilter: "blur(2px)",
                                zIndex: 10,
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                borderRadius: "8px"
                            }}
                        >
                            <PageLoader message={FormType === "EntryForm" ? "Submitting request..." : "Updating request..."} minHeight="100%" />
                        </div>
                    )}
                </div>
            </Panel>
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} type={popupType} />
        </>
    );
};

export default memo(ProjectEntryForm);
