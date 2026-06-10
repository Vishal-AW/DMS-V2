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
    FontIcon
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
import { getParent } from "../../../../Services/FolderMasterService";


// Interface for the nested folder structure
interface FolderNode {
    ID: number;
    FolderName: string;
    IsParentFolder: boolean;
    Active: boolean;
    ParentFolderIdId: number | null; // This will store the ID of the parent folder
    children?: FolderNode[];
    isExpanded?: boolean; // To control the expanded state in the UI
}

// Helper function to build the folder tree
const buildFolderTree = (folders: any[], parentId: number | null = null): FolderNode[] => {
    const tree: FolderNode[] = [];

    // Find direct children of the current parentId
    const directChildren = folders.filter(folder => {
        return folder.ParentFolderIdId === parentId;
    });

    for (const folder of directChildren) {
        const node: FolderNode = {
            ID: folder.ID,
            FolderName: folder.FolderName,
            IsParentFolder: folder.IsParentFolder,
            Active: folder.Active,
            ParentFolderIdId: folder.ParentFolderIdId,
            children: buildFolderTree(folders, folder.ID), // Recursively build children
            isExpanded: true // Root folders must be expanded by default
        };
        tree.push(node);
    }
    return tree;
};

interface FolderTreeNodeProps {
    node: FolderNode;
    depth: number;
}


//proper working code 
// const FolderTreeNode = ({
//     node,
//     depth = 0,
//     isRoot = false,
//     isLast = false,
//     parentLines = []
// }: any): JSX.Element => {
//     const [isExpanded, setIsExpanded] = React.useState(true);
//     const hasChildren = node.children?.length > 0;

//     return (
//         <>
//             <div
//                 style={{
//                     display: "flex",
//                     alignItems: "center",
//                     minHeight: 32,
//                     fontSize: 14,
//                     paddingLeft: depth === 0 ? 8 : depth * 24 + 8,
//                     cursor: hasChildren ? "pointer" : "default",
//                     userSelect: "none",
//                     position: "relative"
//                 }}
//                 onClick={() => hasChildren && setIsExpanded(prev => !prev)}
//             >
//                 {/* Chevron — only for nodes with children, placed BEFORE folder icon */}
//                 <span
//                     style={{
//                         display: "inline-flex",
//                         alignItems: "center",
//                         justifyContent: "center",
//                         width: 16,
//                         marginRight: 4,
//                         fontSize: 13,
//                         color: "#6b7280",
//                         fontWeight: 500,
//                         visibility: hasChildren ? "visible" : "hidden",
//                         transform: isExpanded ? "rotate(90deg)" : "rotate(0deg)",
//                         transition: "transform 0.15s ease"
//                     }}
//                 >
//                     &rsaquo;
//                 </span>

//                 {/* Folder icon */}
//                 <span style={{ marginRight: 6, fontSize: 18, lineHeight: 1 }}>
//                     {hasChildren && isExpanded ? "📂" : "📁"}
//                 </span>

//                 {/* Label */}
//                 <span
//                     style={{
//                         fontWeight: hasChildren ? 600 : 400,
//                         color: "#111827",
//                         fontSize: 14
//                     }}
//                 >
//                     {node.FolderName}
//                 </span>
//             </div>

//             {/* Children */}
//             {hasChildren && isExpanded &&
//                 node.children.map((child: any, index: number) => (
//                     <FolderTreeNode
//                         key={child.ID}
//                         node={child}
//                         depth={depth + 1}
//                         isLast={index === node.children.length - 1}
//                         parentLines={[...parentLines, !isLast]}
//                     />
//                 ))}
//         </>
//     );
// };

const FolderTreeNode = ({
    node,
    depth = 0,
    isLast = false,
    parentLines = []
}: any): JSX.Element => {
    const [isExpanded, setIsExpanded] = React.useState(true);
    const hasChildren = node.children?.length > 0;
    const INDENT = 24; // px per depth level
    const LINE_X = 16; // x offset of vertical line within each indent block

    return (
        <>
            <div
                style={{
                    position: "relative",
                    display: "flex",
                    alignItems: "center",
                    minHeight: 32,
                    fontSize: 14,
                    cursor: hasChildren ? "pointer" : "default",
                    userSelect: "none"
                }}
                onClick={() => hasChildren && setIsExpanded(prev => !prev)}
            >
                {/* ── Ancestor vertical lines ── */}
                {parentLines.map((show: boolean, idx: number) =>
                    show ? (
                        <div
                            key={idx}
                            style={{
                                position: "absolute",
                                left: idx * INDENT + LINE_X,
                                top: 0,
                                bottom: 0,
                                width: 1,
                                background: "#c0c0c0"
                            }}
                        />
                    ) : null
                )}

                {/* ── Current node: vertical + horizontal connector ── */}
                {depth > 0 && (
                    <>
                        {/* Vertical: top → midpoint (stops if last child) */}
                        <div
                            style={{
                                position: "absolute",
                                left: (depth - 1) * INDENT + LINE_X,
                                top: 0,
                                height: isLast ? "50%" : "100%",
                                width: 1,
                                background: "#c0c0c0"
                            }}
                        />
                        {/* Horizontal: midpoint connector to icon */}
                        <div
                            style={{
                                position: "absolute",
                                left: (depth - 1) * INDENT + LINE_X,
                                top: "50%",
                                width: INDENT - LINE_X + 2,
                                height: 1,
                                background: "#c0c0c0"
                            }}
                        />
                    </>
                )}

                {/* ── Content row ── */}
                <div
                    style={{
                        display: "flex",
                        alignItems: "center",
                        paddingLeft: depth * INDENT + 8
                    }}
                >
                    {/* Chevron */}
                    <span
                        style={{
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            width: 14,
                            marginRight: 4,
                            fontSize: 12,
                            color: "#6b7280",
                            visibility: hasChildren ? "visible" : "hidden",
                            transform: isExpanded ? "rotate(90deg)" : "rotate(0deg)",
                            transition: "transform 0.15s ease",
                            lineHeight: 1
                        }}
                    >
                        {/* Using a proper right-pointing chevron */}
                        <svg width="8" height="12" viewBox="0 0 8 12" fill="none">
                            <path d="M1.5 1.5L6.5 6L1.5 10.5" stroke="#6b7280" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>
                    </span>

                    {/* Folder icon */}
                    <span style={{ marginRight: 6, fontSize: 17, lineHeight: 1 }}>
                        {hasChildren && isExpanded ? "📂" : "📁"}
                    </span>

                    {/* Label */}
                    <span
                        style={{
                            fontWeight: hasChildren ? 600 : 400,
                            color: "#111827",
                            fontSize: 14
                        }}
                    >
                        {node.FolderName}
                    </span>
                </div>
            </div>

            {/* ── Children ── */}
            {hasChildren && isExpanded &&
                node.children.map((child: any, index: number) => (
                    <FolderTreeNode
                        key={child.ID}
                        node={child}
                        depth={depth + 1}
                        isLast={index === node.children.length - 1}
                        parentLines={[...parentLines, !isLast]}
                    />
                ))}
        </>
    );
};


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

      // New states for folder tree preview
    const [viewTemplateName, setViewTemplateName] = useState("");
    const [viewFolderTree, setViewFolderTree] = useState<FolderNode[]>([]); // New state
    const [isFolderTreeLoading, setIsFolderTreeLoading] = useState<boolean>(false); // New state


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
        setIsFolderTreeLoading(true); // Start loading for the folder tree

        const templateRes = await getTemplateDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );
        const templateData = templateRes.value[0];
        setViewTemplateName(templateData.Name);
        setTemplate(templateData.Name); // Keep existing state update
        setIsActiveTemplateStatus(templateData.Active); // Keep existing state update

        const allFoldersRes: any = await getParent(context.pageContext.web.absoluteUrl, context.spHttpClient);
        const allFolders = allFoldersRes?.value || [];

        // Filter folders belonging to the selected template
        const filteredTemplateFolders = allFolders.filter(
            (item: any) => item.TemplateName?.Name === templateData.Name
        );

        // Build the nested tree structure
        const nestedTree = buildFolderTree(filteredTemplateFolders, null); // Start with null for root parents
        setViewFolderTree(nestedTree);
        setIsFolderTreeLoading(false); // End loading for the folder tree
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

    // const validation = () => {
    //     if (!Template.trim()) {
    //         setTemplateErr("Template Name is required");
    //         return false;
    //     }
    //     return true;
    // };

    const validation = () => {

        const name = Template.trim().toLowerCase();

        
        if (name !== name?.trim()) {
              // Starting or ending space validation
            setTemplateErr("Spaces are not allowed at starting or ending");
            return false;
        }
        if (!name) {
            setTemplateErr("Template Name is required");
            return false;
        }
          // Special character validation
        else if (!/^[a-zA-Z0-9 ]+$/.test(name)) {
             setTemplateErr("Special characters are not allowed");
            return false;
        }


        // Duplicate validation
        const isDuplicate = tableData.some((item: any) =>
            item.Name?.toLowerCase() === name &&
            item.ID !== TemplateCurrentEditID   // allow same record during edit
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

                    <TextField
                        placeholder="Search..."
                        value={searchText}
                        onChange={(_, val) => setSearchText(val || "")}
                        styles={{ root: { width: 300 } }}
                    />

                    <PrimaryButton
                        text="Add Template"
                        onClick={openTemplatePanel}
                        styles={getAddActionButtonStyles()}
                    />

                </div>

                <ReactTableComponent
                    rowData={filteredData}
                    columnDefs={TemplateTablecolumns}
                />

            </Stack>

            <Panel
                isOpen={isTemplatePanelOpen}
                onDismiss={closeTemplatePanel}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
                // headerText={isTemplateEditMode ? "Edit Template" : "Add Template"}
                 headerText={isTemplateViewMode ? `View ${viewTemplateName} Structure` : (isTemplateEditMode ? "Edit Template" : "Add Template")}
                isFooterAtBottom={true}
                // onRenderFooterContent={() => (
                //     <>
                //         <PrimaryButton
                //             text={isTemplateEditMode ? "Update" : "Submit"}
                //             onClick={SaveItemData}
                //             styles={getPrimaryActionButtonStyles(8)}
                //         />

                //         <DefaultButton
                //             text="Cancel"
                //             onClick={closeTemplatePanel}
                //             styles={getSecondaryActionButtonStyles()}
                //         />
                //     </>
                // )}
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

                        {/* Folder Structure Preview */}
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
                                    // maxHeight: 240,
                                     maxHeight: 650,
                                    overflowY: "auto",
                                    padding: "6px 0"
                                }}
                            >
                                {isFolderTreeLoading ? (
                                    <PageLoader message="Loading folder structure..." />
                                ) : viewFolderTree.length > 0 ? (
                                    viewFolderTree.map((node: any) => (
                                        <FolderTreeNode
                                            key={node.ID}
                                            node={node}
                                            depth={0}
                                            isRoot
                                        />
                                    ))
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
