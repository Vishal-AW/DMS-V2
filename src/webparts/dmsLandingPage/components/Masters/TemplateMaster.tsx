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
import { Link } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "../styles/global.css";
import ReactTableComponent from "../ResuableComponents/ReusableDataTable";
import PopupBox from "../../common/component/PopupBox";
import PageLoader from "../../common/component/PageLoader";

import {
    getTemplate,
    getTemplateDataByID,
    SaveTemplateMaster,
    UpdateTemplateMaster
} from "../../../../Services/TemplateService";

interface ITempletMaster {
    context: WebPartContext;
}

export default function TemplateMaster({ context }: ITempletMaster): JSX.Element {

    const [tableData, setTableData] = useState<any[]>([]);
    const [searchText, setSearchText] = useState("");

    const [isTemplatePanelOpen, setIsTemplatePanelOpen] = useState(false);
    const [isTemplateEditMode, setIsTemplateEditMode] = useState(false);

    const [Template, setTemplate] = useState("");
    const [isActiveTemplateStatus, setIsActiveTemplateStatus] = useState(true);
    const [TemplateErr, setTemplateErr] = useState("");

    const [TemplateCurrentEditID, setTemplateCurrentEditID] = useState<number>(0);

    const [isPopupVisible, setIsPopupVisible] = useState(false);
    const [alertMsg, setAlertMsg] = useState("");
    const [isLoading, setIsLoading] = useState(true);

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

    const clearFields = () => {
        setTemplate("");
        setTemplateErr("");
        setTemplateCurrentEditID(0);
        setIsActiveTemplateStatus(true);
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

        if (!name) {
            setTemplateErr("Template Name is required");
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
                headerText={isTemplateEditMode ? "Edit Template" : "Add Template"}
                isFooterAtBottom={true}
                onRenderFooterContent={() => (
                    <>
                        <PrimaryButton
                            text={isTemplateEditMode ? "Update" : "Submit"}
                            onClick={SaveItemData}
                            styles={{
                                root: {
                                    backgroundColor: "#009ef7",
                                    borderRadius: 6,
                                    marginRight: 10
                                }
                            }}
                        />

                        <DefaultButton
                            text="Cancel"
                            onClick={closeTemplatePanel}
                            styles={{
                                root: { borderRadius: 6 }
                            }}
                        />
                    </>
                )}
            >

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
