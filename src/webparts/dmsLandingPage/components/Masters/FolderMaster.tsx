

import * as React from "react";
import { useEffect, useState } from "react";
import {
    Stack,
    TextField,
    Panel,
    PanelType,
    DefaultButton,
    PrimaryButton,
    Toggle,
    FontIcon
} from "@fluentui/react";
import { Badge, Field } from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";
// import styles from "./Master.module.scss";
import "../styles/global.css";
import ReactTableComponent from "../ResuableComponents/ReusableDataTable";
import PopupBox from "../../common/component/PopupBox";
import PageLoader from "../../common/component/PageLoader";

import {
    getParent,
    getChildDataByID,
    SaveFolderMaster,
    UpdateFolderMaster,
    getTemplateDataByID,

} from "../../../../Services/MasFolderService";
import Select from "react-select";
import { getTemplate } from "../../../../Services/TemplateService";

interface IFolderMaster {
    context: WebPartContext;
}

export default function FolderMaster({ context }: IFolderMaster): JSX.Element {

    const [tableData, setTableData] = useState<any[]>([]);
    const [searchText, setSearchText] = useState("");

    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [isEditMode, setIsEditMode] = useState(false);

    const [FolderName, setFolderName] = useState("");
    const [IsChildFolder, setIsChildFolder] = useState(false);
    const [Active, setActive] = useState(true);

    const [TemplateList, setTemplateList] = useState<any[]>([]);
    const [ParentFolderList, setParentFolderList] = useState<any[]>([]);

    const [TemplateNameId, setTemplateNameId] = useState<number | undefined>();
    const [ParentFolderIdId, setParentFolderIdId] = useState<number | undefined>();

    const [currentEditID, setCurrentEditID] = useState<number>(0);

    const [nameError, setNameError] = useState("");
    const [templateError, setTemplateError] = useState("");
    const [parentError, setParentError] = useState("");

    const [isPopupVisible, setIsPopupVisible] = useState(false);
    const [alertMsg, setAlertMsg] = useState("");
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        Promise.all([fetchData(), fetchTemplates()]).finally(() => setIsLoading(false));
    }, []);

    const fetchData = async () => {

        const res: any = await getParent(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient
        );

        setTableData(res?.value || []);

    };

    const fetchTemplates = async () => {

        const res: any = await getTemplate(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient
        );

        setTemplateList(res?.value || []);

    };

    // When template is selected in the panel, fetch parent folders for that template
    // getTemplateDataByID returns folders filtered by TemplateId where IsParentFolder = false
    // (i.e. root/parent folders that can be selected as a parent for child folders)
    const loadParentFoldersForTemplate = async (templateId: number) => {

        const res: any = await getTemplateDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            templateId
        );

        setParentFolderList(res?.value || []);

    };

    const handleTemplateChange = async (templateId: number) => {

        setTemplateNameId(templateId);
        setParentFolderIdId(undefined);
        setParentFolderList([]);
        setTemplateError("");

        if (templateId) {
            await loadParentFoldersForTemplate(templateId);
        }

    };

    const filteredData = tableData.filter((item) =>
        item.FolderName?.toLowerCase().includes(searchText.toLowerCase())
    );

    const clearFields = () => {

        setFolderName("");
        setIsChildFolder(false);
        setActive(true);
        setTemplateNameId(undefined);
        setParentFolderIdId(undefined);
        setParentFolderList([]);
        setCurrentEditID(0);
        setNameError("");
        setTemplateError("");
        setParentError("");

    };

    const openPanel = () => {

        clearFields();
        setIsEditMode(false);
        setIsPanelOpen(true);

    };

    const openEditPanel = async (id: number) => {

        clearFields();
        setIsEditMode(true);
        setIsPanelOpen(true);

        const res: any = await getChildDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );

        const data = res.value[0];

        setCurrentEditID(data.ID);
        setFolderName(data.FolderName);
        setActive(data.Active);
        setTemplateNameId(data.TemplateName?.ID);

        // Load parent folder list for this template so dropdown is populated
        if (data.TemplateName?.ID) {
            await loadParentFoldersForTemplate(data.TemplateName.ID);
        }


        if (data.IsParentFolder === true && data.ParentFolderId) {

            setIsChildFolder(true);
            // setParentFolderIdId(data.ParentFolderId.ID);
            setParentFolderIdId(data.ParentFolderId.ID || data.ParentFolderId.Id);

        } else {

            setIsChildFolder(false);
            setParentFolderIdId(undefined);

        }

    };

    const closePanel = () => {

        clearFields();
        setIsPanelOpen(false);

    };

    const clearErrors = () => {

        setNameError("");
        setTemplateError("");
        setParentError("");

    };

    const validation = (): boolean => {

        clearErrors();
        let isValid = true;

        if (!FolderName.trim()) {
            setNameError("This field is required");
            isValid = false;
        }

        if (!TemplateNameId) {
            setTemplateError("This field is required");
            isValid = false;
        }

        if (IsChildFolder && !ParentFolderIdId) {
            setParentError("This field is required");
            isValid = false;
        }

        if (!isValid) return false;

        const currentFolderName = FolderName.trim().toLowerCase();
        const currentTemplateId = Number(TemplateNameId);
        const currentParentId = IsChildFolder ? Number(ParentFolderIdId) : null;

        const isDuplicate = tableData.some((data) => {

            if (isEditMode && data.ID === currentEditID) return false;

            const dataFolderName = (data.FolderName || "").trim().toLowerCase();
            const dataTemplateId = Number(data.TemplateName?.ID);
            const dataParentId = data.ParentFolderIdId
                ? Number(data.ParentFolderIdId)
                : data.ParentFolderId?.ID
                    ? Number(data.ParentFolderId.ID)
                    : null;

            return (
                dataFolderName === currentFolderName &&
                dataTemplateId === currentTemplateId &&
                dataParentId === currentParentId
            );

        });

        if (isDuplicate) {
            setNameError("This combination of Folder Name, Template, and Parent Folder already exists.");
            return false;
        }

        return true;

    };


    const SaveItemData = async () => {

        if (!validation()) return;

        try {

            // Reference file logic for IsParentFolder field:
            // IsChildFolder toggle ON  → IsParentFolder = true  (it's a child, has a parent)
            // IsChildFolder toggle OFF → IsParentFolder = false (it's a root/parent folder)
            const option: any = {
                FolderName: FolderName.trim(),
                Active: Active,
                TemplateNameId: TemplateNameId,
                IsParentFolder: IsChildFolder,
                ParentFolderIdId: IsChildFolder ? (ParentFolderIdId ?? null) : null
            };

            if (!isEditMode) {

                await SaveFolderMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option
                );

                setAlertMsg("Folder Added Successfully");

            } else {

                await UpdateFolderMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option,
                    currentEditID
                );

                setAlertMsg("Folder Updated Successfully");

            }

            setIsPopupVisible(true);
            setIsPanelOpen(false);
            fetchData();

        } catch (error) {

            console.error("Save Error:", error);

        }

    };

    const hidePopup = () => {

        setIsPopupVisible(false);

    };

    const TemplateOptions = TemplateList.map((item: any) => ({
        key: item.ID,
        text: item.Name
    }));

    // Parent folder dropdown options — from dynamically loaded list per template
    const ParentFolderOptions = ParentFolderList.map((item: any) => ({
        key: item.ID,
        text: item.FolderName
    }));

    const FolderColumns = [

        { headerName: "Sr No", valueGetter: "node.rowIndex + 1", width: 90 },

        { headerName: "Folder Name", field: "FolderName" },

        {
            headerName: "Template",
            valueGetter: (params: any) => params.data.TemplateName?.Name
        },

        {
            headerName: "Parent Folder",
            valueGetter: (params: any) => params.data.ParentFolderId?.FolderName
        },

        {
            headerName: "Active",
            field: "Active",
            cellRenderer: (params: any) => {

                const isActive = params.value;

                return (
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <Badge appearance="filled" color={isActive ? "success" : "informative"} />
                        {isActive ? "Active" : "Inactive"}
                    </div>
                );

            }
        },

        {
            headerName: "Action",
            cellRenderer: (params: any) => (

                <FontIcon
                    iconName="EditSolid12"
                    style={{
                        color: "#009ef7",
                        cursor: "pointer",
                        backgroundColor: "#f5f8fa",
                        padding: "7px 10px",
                        borderRadius: "6px"
                    }}
                    onClick={() => openEditPanel(params.data.ID)}
                />

            ),
            width: 120
        }
    ];

    if (isLoading) {
        return <PageLoader message="Loading folder master..." minHeight="72vh" />;
    }

    return (

        <div>

            <Stack>

                <div style={{ display: "flex", justifyContent: "space-between", padding: 20 }}>

                    <TextField
                        placeholder="Search..."
                        value={searchText}
                        onChange={(_, val) => setSearchText(val || "")}
                        styles={{ root: { width: 300 } }}
                    />

                    <PrimaryButton text="Add Folder" onClick={openPanel} />

                </div>

                <ReactTableComponent
                    rowData={filteredData}
                    columnDefs={FolderColumns}
                />

            </Stack>

            <Panel
                isOpen={isPanelOpen}
                onDismiss={closePanel}
                type={PanelType.medium}
                headerText={isEditMode ? "Edit Folder" : "Add Folder"}
                isFooterAtBottom
                onRenderFooterContent={() => (

                    <>
                        <PrimaryButton text={isEditMode ? "Update" : "Submit"} onClick={SaveItemData} />
                        <DefaultButton text="Cancel" onClick={closePanel} />
                    </>

                )}
            >

                <Field>

                    <label className="Headerlabel">Folder Name <span style={{ color: "red" }}>*</span></label>

                    <TextField
                        value={FolderName}
                        onChange={(_, val) => {
                            setFolderName(val || "");
                            setNameError("");
                        }}
                    />

                    {nameError && (
                        <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>
                            {nameError}
                        </p>
                    )}
                </Field>


                <Field>
                    <label className="Headerlabel">Is Child Folder?</label>
                    <Toggle
                        checked={IsChildFolder}
                        onChange={(_, val) => {
                            setIsChildFolder(!!val);
                            if (!val) setParentFolderIdId(undefined);
                            setParentError("");
                        }}
                    />
                </Field>

                <Field>

                    <label className="Headerlabel">Template <span style={{ color: "red" }}>*</span></label>

                    <Select
                        options={TemplateOptions}
                        value={TemplateOptions.find((opt) => opt.key === TemplateNameId)}
                        onChange={(selected: any) => {
                            if (selected) {
                                handleTemplateChange(selected.key);
                            } else {
                                setTemplateNameId(undefined);
                                setParentFolderIdId(undefined);
                                setParentFolderList([]);
                            }
                        }}
                        placeholder="Select Template"
                        getOptionLabel={(e: any) => e.text}
                        getOptionValue={(e: any) => String(e.key)}
                    />
                    {templateError && (
                        <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>
                            {templateError}
                        </p>
                    )}
                </Field>


                {IsChildFolder && (

                    <Field>

                        <label className="Headerlabel">Parent Folder <span style={{ color: "red" }}>*</span></label>


                        <Select
                            options={ParentFolderOptions}
                            value={ParentFolderOptions.find((opt) => opt.key === ParentFolderIdId)}
                            onChange={(selected: any) => {
                                setParentFolderIdId(selected?.key);
                                setParentError("");
                            }}
                            placeholder="Select Parent Folder"
                            getOptionLabel={(e: any) => e.text}
                            getOptionValue={(e: any) => String(e.key)}
                        />

                        {parentError && (
                            <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>
                                {parentError}
                            </p>
                        )}
                    </Field>

                )}


                <Field>
                    <label className="Headerlabel">Active</label>
                    <Toggle
                        checked={Active}
                        onChange={(_, val) => setActive(!!val)}
                    />
                </Field>
            </Panel>

            <PopupBox
                isPopupBoxVisible={isPopupVisible}
                hidePopup={hidePopup}
                msg={alertMsg}
                type={isEditMode ? "update" : "insert"}
            />

        </div>
    );
}
