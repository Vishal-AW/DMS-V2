/* eslint-disable */
import * as React from "react"; // Keep this import at the top
import { useEffect, useState } from "react";
import {
    Stack,
    TextField,
    Panel,
    PanelType,
    DefaultButton,
    PrimaryButton,
    Toggle,
    FontIcon,
    ChoiceGroup
} from "@fluentui/react";
import { Badge, Field } from "@fluentui/react-components";
import { Link } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../styles/global.css";
import ReactTableComponent from "../ResuableComponents/ReusableDataTable";
import { getAddActionButtonStyles, getPrimaryActionButtonStyles, getSecondaryActionButtonStyles } from "../../common/component/buttonStyles";
import PopupBox from "../../common/component/PopupBox";
import PageLoader from "../../common/component/PageLoader";

import {
    getTemplate,
    getTemplateDataByID,
    SaveTemplateMaster,
    UpdateTemplateMaster
} from "../../../../Services/TemplateService";
import { getParent, getFoldersByTemplateId, UpdateFolderMaster } from "../../../../Services/FolderMasterService";

import TemplateAccordion from "./TemplateAccordion";
import FolderTreeView, { buildFolderTree, FolderNode } from "./FolderTreeView";

// Interface for the nested folder structure
interface FolderNodeLocal {
    ID: number;
    FolderName: string;
    IsParentFolder: boolean;
    Active: boolean;
    ParentFolderIdId: number | null;
    children?: FolderNodeLocal[];
    isExpanded?: boolean;
}

interface ITempletMaster {
    context: WebPartContext;
}

export default function TemplateMaster({ context }: ITempletMaster): JSX.Element {

    const [tableData, setTableData] = useState<any[]>([]);
    const [searchText, setSearchText] = useState("");

    const [isTemplatePanelOpen, setIsTemplatePanelOpen] = useState(false);
    const [isTemplateEditMode, setIsTemplateEditMode] = useState(false);
    //View
    const [isTemplateViewMode, setIsTemplateViewMode] = useState(false);

    const [Template, setTemplate] = useState("");
    const [isActiveTemplateStatus, setIsActiveTemplateStatus] = useState(true);
    const [TemplateErr, setTemplateErr] = useState("");

    const [TemplateCurrentEditID, setTemplateCurrentEditID] = useState<number>(0);

    const [isPopupVisible, setIsPopupVisible] = useState(false);
    const [alertMsg, setAlertMsg] = useState("");
    const [isLoading, setIsLoading] = useState(true);

    // New states for folder tree preview in View panel
    const [viewTemplateName, setViewTemplateName] = useState("");
    const [viewFolderTree, setViewFolderTree] = useState<FolderNode[]>([]);
    const [isFolderTreeLoading, setIsFolderTreeLoading] = useState<boolean>(false);

    // View mode toggle: "table" or "accordion"
    const [viewMode, setViewMode] = useState<"table" | "accordion">("accordion");

    useEffect(() => {
        fetchData();
    }, []);

    const fetchData = async () => {
        const res: any = await getTemplate(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient
        );

        setTableData(res?.value || []);
        setIsLoading(false);
    };

    const filteredData = tableData.filter((item) =>
        item.Name?.toLowerCase().includes(searchText.toLowerCase())
    );

    const openTemplatePanel = () => {
        clearFields();
        setIsTemplateEditMode(false);
        setIsTemplatePanelOpen(true);
    };

    const openEditTemplatePanel = async (id: number) => {
        clearFields();

        setIsTemplateEditMode(true);
        setIsTemplatePanelOpen(true);

        const res = await getTemplateDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );

        const data = res.value[0];

        setTemplateCurrentEditID(data.ID);
        setTemplate(data.Name);
        setIsActiveTemplateStatus(data.Active);
    };

    const openViewTemplatePanel = async (id: number) => {
        clearFields();

        setIsTemplateViewMode(true);
        setIsTemplateEditMode(false);
        setIsTemplatePanelOpen(true);
        setIsFolderTreeLoading(true);

        const templateRes = await getTemplateDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );
        const templateData = templateRes.value[0];
        setViewTemplateName(templateData.Name);
        setTemplate(templateData.Name);
        setIsActiveTemplateStatus(templateData.Active);

        // Use the new lazy-loading method for folders by template ID
        const foldersRes: any = await getFoldersByTemplateId(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );
        const allFolders = foldersRes?.value || [];

        // Build the nested tree structure
        const nestedTree = buildFolderTree(allFolders, null);
        setViewFolderTree(nestedTree);
        setIsFolderTreeLoading(false);
    };

    const clearFields = () => {
        setTemplate("");
        setTemplateErr("");
        setTemplateCurrentEditID(0);
        setIsActiveTemplateStatus(true);
        setIsTemplateViewMode(false);
    };

    const closeTemplatePanel = () => {
        clearFields();
        setIsTemplatePanelOpen(false);
    };

    const validation = () => {

        const name = Template.trim().toLowerCase();

        if (name !== name?.trim()) {
            setTemplateErr("Spaces are not allowed at starting or ending");
            return false;
        }
        if (!name) {
            setTemplateErr("Template Name is required");
            return false;
        } else if (!/^[a-zA-Z0-9 ]+$/.test(name)) {
            setTemplateErr("Special characters are not allowed");
            return false;
        }

        // Duplicate validation
        const isDuplicate = tableData.some((item: any) =>
            item.Name?.toLowerCase() === name &&
            item.ID !== TemplateCurrentEditID
        );

        if (isDuplicate) {
            setTemplateErr("This template name already exists");
            return false;
        }

        setTemplateErr("");
        return true;
    };

    const SaveItemData = async () => {
        if (!validation()) return;

        let option = {
            __metadata: { type: "SP.Data.DMS_x005f_TemplateListItem" },
            Name: Template.trim(),
            Active: isActiveTemplateStatus
        };

        try {
            if (!isTemplateEditMode) {
                await SaveTemplateMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option
                );
                setAlertMsg("Template Added Successfully");
            } else {
                await UpdateTemplateMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option,
                    TemplateCurrentEditID
                );
                setAlertMsg("Template Updated Successfully");
            }

            setIsPopupVisible(true);
            setIsTemplatePanelOpen(false);
            fetchData();

        } catch (error) {
            console.error("Save Error:", error);
        }
    };

    const hidePopup = () => {
        setIsPopupVisible(false);
    };

    const TemplateTablecolumns = [
        {
            headerName: "Sr No",
            valueGetter: "node.rowIndex + 1",
            width: 90
        },
        {
            headerName: "Template Name",
            field: "Name"
        },
        {
            headerName: "Active",
            field: "Active",
            cellRenderer: (params: any) => {
                const isActive = params.value;

                return (
                    <div
                        style={{
                            display: "flex",
                            alignItems: "center",
                            gap: "8px"
                        }}
                    >
                        <Badge
                            appearance="filled"
                            color={isActive ? "success" : "informative"}
                        />
                        {isActive ? "Active" : "Inactive"}
                    </div>
                );
            }
        },
        {
            headerName: "Action",
            cellRenderer: (params: any) => (
                <div style={{ display: "flex", gap: "8px" }}>

                    <FontIcon
                        iconName="EditSolid12"
                        style={{
                            color: "#009ef7",
                            cursor: "pointer",
                            backgroundColor: "#f5f8fa",
                            padding: "7px 10px",
                            borderRadius: "6px"
                        }}
                        onClick={() => openEditTemplatePanel(params.data.ID)}
                    />

                    <FontIcon
                        iconName="RedEye"
                        style={{
                            color: "#009ef7",
                            cursor: "pointer",
                            backgroundColor: "#f5f8fa",
                            padding: "7px 10px",
                            borderRadius: "6px"
                        }}
                        onClick={() => openViewTemplatePanel(params.data.ID)}
                    />
                </div>
            ),
            width: 120
        }
    ];

    if (isLoading) {
        return <PageLoader message="Loading templates..." minHeight="72vh" />;
    }

    return (
        <div>

            {/* Breadcrumb */}
            <nav
                style={{
                    padding: "14px 22px",
                    background: "#ffffff",
                    borderBottom: "1px solid #e4e6ef"
                }}
            >
                <ol
                    style={{
                        display: "flex",
                        listStyle: "none",
                        margin: 0,
                        padding: 0,
                        fontSize: "14px"
                    }}
                >
                    <li style={{ marginRight: 8 }}>
                        <Link to="/" style={{ textDecoration: "none", color: "#181c32" }}>
                            Dashboard
                        </Link>
                    </li>

                    <li style={{ marginRight: 8 }}>/</li>

                    <li style={{ color: "#009ef7", fontWeight: 600 }}>
                        Template Master
                    </li>
                </ol>
            </nav>

            <Stack>

                <div style={{ display: "flex", justifyContent: "space-between", padding: 20 }}>

                    <div style={{ display: "flex", gap: 12, alignItems: "center" }}>
                        <TextField
                            placeholder="Search..."
                            value={searchText}
                            onChange={(_, val) => setSearchText(val || "")}
                            styles={{ root: { width: 300 } }}
                        />

                        {/* View Mode Toggle */}
                        <ChoiceGroup
                            selectedKey={viewMode}
                            options={[
                                { key: "accordion", text: "Accordion View", iconProps: { iconName: "BulletedList" } },
                                { key: "table", text: "Table View", iconProps: { iconName: "GridViewSmall" } }
                            ]}
                            onChange={(_, option) => {
                                if (option) setViewMode(option.key as "table" | "accordion");
                            }}
                            styles={{
                                root: { display: "flex" },
                                flexContainer: { display: "flex", gap: 8 }
                            }}
                        />
                    </div>

                    <PrimaryButton
                        text="Add Template"
                        onClick={openTemplatePanel}
                        styles={getAddActionButtonStyles()}
                    />

                </div>

                {/* Conditional Rendering: Accordion View or Table View */}
                {viewMode === "accordion" ? (
                    <div style={{ padding: "0 20px 20px" }}>
                        <TemplateAccordion
                            context={context}
                            templates={filteredData}
                            onRefreshTemplates={fetchData}
                        />
                    </div>
                ) : (
                    <ReactTableComponent
                        rowData={filteredData}
                        columnDefs={TemplateTablecolumns}
                    />
                )}

            </Stack>

            <Panel
                isOpen={isTemplatePanelOpen}
                onDismiss={closeTemplatePanel}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                headerText={isTemplateViewMode ? `View ${viewTemplateName} Structure` : (isTemplateEditMode ? "Edit Template" : "Add Template")}
                isFooterAtBottom={true}
                onRenderFooterContent={() => (
                    <>
                        {!isTemplateViewMode && (
                            <PrimaryButton
                                text={isTemplateEditMode ? "Update" : "Submit"}
                                onClick={SaveItemData}
                                styles={getPrimaryActionButtonStyles(8)}
                            />
                        )}
                        <DefaultButton
                            text="Cancel"
                            onClick={closeTemplatePanel}
                            styles={getSecondaryActionButtonStyles()}
                        />
                    </>
                )}
            >

                {isTemplateViewMode ? (
                    <div style={{ padding: "0 4px" }}>
                        {/* Template Name */}
                        <div style={{ marginBottom: 16 }}>
                            <label
                                style={{
                                    display: "block",
                                    fontSize: 11,
                                    fontWeight: 600,
                                    letterSpacing: "0.07em",
                                    color: "#6b7280",
                                    textTransform: "uppercase",
                                    marginBottom: 6
                                }}
                            >
                                Template Name
                            </label>
                            <TextField
                                value={viewTemplateName}
                                disabled
                                readOnly
                                styles={{
                                    root: { width: "100%" },
                                    field: {
                                        background: "#f3f4f6",
                                        border: "1px solid #e5e7eb",
                                        borderRadius: 6,
                                        color: "#374151",
                                        fontSize: 14
                                    }
                                }}
                            />
                        </div>

                        {/* Folder Structure Preview - Enhanced with drag-drop enabled tree */}
                        <div>
                            <label
                                style={{
                                    display: "block",
                                    fontSize: 11,
                                    fontWeight: 600,
                                    letterSpacing: "0.07em",
                                    color: "#6b7280",
                                    textTransform: "uppercase",
                                    marginBottom: 8
                                }}
                            >
                                Folder Structure Preview
                            </label>

                            <div
                                style={{
                                    border: "1px solid #e5e7eb",
                                    borderRadius: 8,
                                    background: "#f9fafb",
                                    maxHeight: 650,
                                    overflowY: "auto",
                                    padding: "6px 0"
                                }}
                            >
                                {isFolderTreeLoading ? (
                                    <PageLoader message="Loading folder structure..." />
                                ) : viewFolderTree.length > 0 ? (
                                    <FolderTreeView
                                        folders={viewFolderTree}
                                        templateId={TemplateCurrentEditID}
                                        templateName={viewTemplateName}
                                        onRefreshTree={async () => {
                                            // Refresh the view tree
                                            const foldersRes: any = await getFoldersByTemplateId(
                                                context.pageContext.web.absoluteUrl,
                                                context.spHttpClient,
                                                TemplateCurrentEditID
                                            );
                                            const allFolders = foldersRes?.value || [];
                                            const nestedTree = buildFolderTree(allFolders, null);
                                            setViewFolderTree(nestedTree);
                                        }}
                                        onAddFolder={() => { }}
                                        onDragEnd={async (draggedFolderId, newParentId, newOrder) => {
                                            // Update folder after drag in view mode
                                            const option: any = {
                                                ParentFolderIdId: newParentId
                                            };
                                            await UpdateFolderMaster(
                                                context.pageContext.web.absoluteUrl,
                                                context.spHttpClient,
                                                option,
                                                draggedFolderId
                                            );
                                            // Refresh the view tree
                                            const foldersRes: any = await getFoldersByTemplateId(
                                                context.pageContext.web.absoluteUrl,
                                                context.spHttpClient,
                                                TemplateCurrentEditID
                                            );
                                            const allFolders = foldersRes?.value || [];
                                            const nestedTree = buildFolderTree(allFolders, null);
                                            setViewFolderTree(nestedTree);
                                        }}
                                    />
                                ) : (
                                    <div
                                        style={{
                                            padding: "8px 12px",
                                            fontSize: 13,
                                            color: "#6b7280"
                                        }}
                                    >
                                        No folders configured for this template.
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>
                ) : (
                    <>
                        <Field>
                            <label className="Headerlabel">Template Name <span style={{ color: "red" }}>*</span></label>
                            <TextField
                                value={Template}
                                onChange={(_, val) => {
                                    setTemplate(val || "");
                                    setTemplateErr("");
                                }}
                                errorMessage={TemplateErr}
                                placeholder="Enter Template Name"
                            />
                        </Field>
                        <Field>
                            <label className="Headerlabel">Active Status</label>
                            <Toggle
                                checked={isActiveTemplateStatus}
                                onChange={(_, checked) => setIsActiveTemplateStatus(!!checked)}
                            />
                        </Field>
                    </>
                )}

            </Panel>

            <PopupBox
                isPopupBoxVisible={isPopupVisible}
                hidePopup={hidePopup}
                msg={alertMsg}
                type={isTemplateEditMode ? "update" : "insert"}
            />

        </div>
    );
}
