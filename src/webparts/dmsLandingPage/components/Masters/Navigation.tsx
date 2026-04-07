/* eslint-disable */
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
import { Badge } from "@fluentui/react-components";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Select from "react-select";
import PopupBox from "../../common/component/PopupBox";
import ReactTableComponent from "../ResuableComponents/ReusableDataTable";
import {
    SaveNavigationMaster,
    getdata,
    getDataByID,
    UpdateNavigationMaster
} from "../../../../Services/NavigationService";
import { getUserIdFromLoginName } from "../../../../DAL/Commonfile";
import {
    PeoplePicker,
    PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Field } from "@fluentui/react-components";
import { Link } from "react-router-dom";

interface INavigation {
    context: WebPartContext;
}

const ORDER_OPTIONS = Array.from({ length: 20 }, (_, i) => ({
    value: i + 1,
    label: String(i + 1)
}));

export default function Navigation({ context }: INavigation) {
    const [tableData, setTableData] = useState<any[]>([]);
    const [searchText, setSearchText] = useState("");

    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [isEditMode, setIsEditMode] = useState(false);

    const [MenuName, setMenuName] = useState("");
    const [URL, setURL] = useState("");
    const [isActive, setIsActive] = useState(true);
    const [isNextActive, setIsNextActive] = useState(true);
    const [isParentMenu, setIsParentMenu] = useState(true);
    const [OrdeID, setOrdeID] = useState<any>(null);
    const [ParentMenuID, setParentMenuID] = useState<any>(null);
    const [parentOptions, setParentOptions] = useState<any[]>([]);

    const [selectedUsers, setSelectedUsers] = useState<any[]>([]);
    const [assignID, setAssignID] = useState<string[]>([]);
    const [assignEmail, setAssignEmail] = useState<string[]>([]);

    const [currentEditID, setCurrentEditID] = useState<number>(0);

    const [isPopupVisible, setIsPopupVisible] = useState(false);
    const [alertMsg, setAlertMsg] = useState("");

    // Errors
    const [MenuNameErr, setMenuNameErr] = useState("");
    const [URLErr, setURLErr] = useState("");
    const [ParentMenuIDErr, setParentMenuIDErr] = useState("");
    const [OrdeIDErr, setOrdeIDErr] = useState("");
    const [AccessErr, setAccessErr] = useState("");

    useEffect(() => {
        fetchData();
    }, []);

    const fetchData = async () => {
        const res: any = await getdata(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient
        );
        const data = res?.value || [];
        setTableData(data);
        setParentOptions(
            data.map((item: any) => ({ value: item.ID, label: item.MenuName }))
        );
    };


    const peoplePickerContext: any = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory,
        spHttpClient: context.spHttpClient
    };

    const onPeoplePickerChange = (items: any[]) => {
        setSelectedUsers(items);
        setAssignID(items.map((i) => i.id));
        setAssignEmail(items.map((i) => i.secondaryText));
        if (items.length > 0) setAccessErr("");
    };

    const clearFields = () => {
        setMenuName("");
        setURL("");
        setIsActive(true);
        setIsNextActive(true);
        setIsParentMenu(true);
        setParentMenuID(null);
        setOrdeID(null);
        setAssignID([]);
        setAssignEmail([]);
        setSelectedUsers([]);
        setCurrentEditID(0);
        setMenuNameErr("");
        setURLErr("");
        setParentMenuIDErr("");
        setOrdeIDErr("");
        setAccessErr("");
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

        const res: any = await getDataByID(
            context.pageContext.web.absoluteUrl,
            context.spHttpClient,
            id
        );
        const data = res.value[0];

        setCurrentEditID(data.ID);
        setMenuName(data.MenuName);
        setURL(data.URL);
        setIsActive(data.Active);
        setIsNextActive(data.Next_Tab);
        setIsParentMenu(data.isParentMenu);
        setOrdeID({ value: data.OrderNo, label: String(data.OrderNo) });

        if (!data.isParentMenu && data.ParentMenuId) {
            setParentMenuID({
                value: data.ParentMenuId?.Id,
                label: data.ParentMenuId?.MenuName
            });
        }

        if (data.Permission && data.Permission.length > 0) {
            const emails = data.Permission.map((p: any) => {
                const parts = (p.Name || "").split('|');
                return parts.length > 1 ? parts[parts.length - 1] : p.Name;
            }).filter(Boolean);

            setAssignEmail(emails);
            setAssignID(emails);
        }
    };

    const validation = (): boolean => {
        let valid = true;

        if (!MenuName.trim()) { setMenuNameErr("Required"); valid = false; }
        if (!URL.trim()) { setURLErr("Required"); valid = false; }
        if (!OrdeID) { setOrdeIDErr("Required"); valid = false; }
        if (!isParentMenu && !ParentMenuID) { setParentMenuIDErr("Required"); valid = false; }
        if (assignID.length === 0) { setAccessErr("Required"); valid = false; }

        return valid;
    };

    const SaveItemData = async () => {
        if (!validation()) return;

        const userIds = await Promise.all(
            assignID.map(async (person: any) => {
                const user = await getUserIdFromLoginName(context, person);
                return user.Id;
            })
        );


        const option = {
            __metadata: { type: "SP.Data.GEN_x005f_NavigationListItem" },
            MenuName: MenuName.trim(),
            PermissionId: { results: userIds },
            ParentMenuIdId: isParentMenu ? null : Number(ParentMenuID?.value),
            URL: URL.trim(),
            Active: Boolean(isActive),
            Next_Tab: Boolean(isNextActive),
            OrderNo: Number(OrdeID?.value),
            isParentMenu: Boolean(isParentMenu)
        };

        if (!isEditMode) {
            await SaveNavigationMaster(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                option
            );
            setAlertMsg("Saved Successfully");
        } else {
            await UpdateNavigationMaster(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                option,
                currentEditID
            );
            setAlertMsg("Updated Successfully");
        }

        setIsPopupVisible(true);
        setIsPanelOpen(false);
        fetchData();
    };


    const columns = [
        { headerName: "Sr No", valueGetter: "node.rowIndex + 1", width: 90 },
        { headerName: "Menu Name", field: "MenuName" },
        { headerName: "URL", field: "URL" },
        {
            headerName: "Active",
            field: "Active",
            cellRenderer: (params: any) => (
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <Badge appearance="filled" color={params.value ? "success" : "informative"} />
                    {params.value ? "Active" : "Inactive"}
                </div>
            )
        },
        {
            headerName: "Next Tab",
            field: "Next_Tab",
            valueGetter: (p: any) => (p.data.Next_Tab ? "Yes" : "No")
        },
        {
            headerName: "Parent Menu",
            valueGetter: (p: any) => p.data.ParentMenuId?.MenuName || "-"
        },
        { headerName: "Order", field: "OrderNo", width: 100 },
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

    return (
        <div>
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
                        Navigation Master
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
                    <PrimaryButton text="Add Menu" onClick={openPanel} />
                </div>


                <ReactTableComponent
                    rowData={tableData}
                    columnDefs={columns}
                    searchText={searchText}
                />
            </Stack>

            <Panel
                isOpen={isPanelOpen}
                onDismiss={() => setIsPanelOpen(false)}
                type={PanelType.large}
                headerText={isEditMode ? "Edit Menu" : "Add Menu"}
                isFooterAtBottom
                onRenderFooterContent={() => (
                    <>
                        <PrimaryButton text={isEditMode ? "Update" : "Submit"} onClick={SaveItemData} />
                        <DefaultButton
                            text="Cancel"
                            onClick={() => setIsPanelOpen(false)}
                            style={{ marginLeft: 8 }}
                        />
                    </>
                )}
            >
                <div className="grid">

                    <div style={{ display: "flex", gap: 16, width: "100%" }}>

                        <div style={{ flex: 1 }}>
                            <Field>
                                <label className="Headerlabel">Menu Name <span style={{ color: "red" }}>*</span></label>
                                <TextField
                                    value={MenuName}
                                    onChange={(_, v) => { setMenuName(v || ""); setMenuNameErr(""); }}
                                    errorMessage={MenuNameErr}
                                />
                            </Field>
                        </div>

                        <div style={{ flex: 1 }}>
                            <Field>
                                <label className="Headerlabel">URL <span style={{ color: "red" }}>*</span></label>
                                <TextField
                                    value={URL}
                                    onChange={(_, v) => { setURL(v || ""); setURLErr(""); }}
                                    errorMessage={URLErr}
                                />
                            </Field>
                        </div>

                        <div style={{ flex: 1 }}>
                            <Field>
                                <label className="Headerlabel">Order <span style={{ color: "red" }}>*</span></label>
                                <Select
                                    options={ORDER_OPTIONS}
                                    value={OrdeID}
                                    onChange={(val: any) => { setOrdeID(val); setOrdeIDErr(""); }}
                                    placeholder="Select Order"
                                />
                                {OrdeIDErr && <p style={{ color: "red", fontSize: 12 }}>{OrdeIDErr}</p>}
                            </Field>
                        </div>

                    </div>


                    <div style={{ display: "flex", gap: 16, marginTop: 16 }}>

                        <div style={{ flex: 1 }}>
                            <label className="Headerlabel">Active</label>
                            <Toggle checked={isActive} onChange={(_, v) => setIsActive(!!v)} />
                        </div>

                        <div style={{ flex: 1 }}>
                            <label className="Headerlabel">Next Tab</label>
                            <Toggle checked={isNextActive} onChange={(_, v) => setIsNextActive(!!v)} />
                        </div>

                        <div style={{ flex: 1 }}>
                            <label className="Headerlabel">Is Parent Menu</label>
                            <Toggle
                                checked={isParentMenu}
                                onChange={(_, v) => {
                                    setIsParentMenu(!!v);
                                    if (!!v) setParentMenuID(null);
                                }}
                            />
                        </div>

                    </div>


                    {!isParentMenu && (
                        <div style={{ marginTop: 16 }}>
                            <Field>
                                <label className="Headerlabel">Parent Menu <span style={{ color: "red" }}>*</span></label>
                                <Select
                                    options={parentOptions}
                                    value={ParentMenuID}
                                    onChange={(val: any) => { setParentMenuID(val); setParentMenuIDErr(""); }}
                                    placeholder="Select Parent Menu"
                                />
                                {ParentMenuIDErr && (
                                    <p style={{ color: "red", fontSize: 12 }}>{ParentMenuIDErr}</p>
                                )}
                            </Field>
                        </div>
                    )}



                    <div style={{ display: "flex", gap: 16, marginTop: 16 }}>
                        <div className="col12">
                            <label className="Headerlabel">Access Permission <span style={{ color: "red" }}>*</span></label>
                            <PeoplePicker
                                context={peoplePickerContext}
                                personSelectionLimit={10}
                                onChange={onPeoplePickerChange}
                                principalTypes={[PrincipalType.User]}
                                defaultSelectedUsers={isEditMode ? assignEmail : []}
                            />
                            {AccessErr && (
                                <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>{AccessErr}</p>
                            )}
                        </div>
                    </div>

                </div>
            </Panel>

            <PopupBox
                isPopupBoxVisible={isPopupVisible}
                hidePopup={() => setIsPopupVisible(false)}
                msg={alertMsg}
            />
        </div>
    );
}