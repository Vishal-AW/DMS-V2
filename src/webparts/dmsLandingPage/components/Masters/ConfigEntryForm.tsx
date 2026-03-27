import {
    DefaultButton,
    FontIcon,
    Panel,
    PanelType,
    PrimaryButton,
    Stack,
    TextField,
    Toggle
} from "@fluentui/react";
import * as React from "react";
import { useEffect, useRef, useState } from "react";
import { SPHttpClient } from "@microsoft/sp-http-base";
import PopupBox from "../../common/component/PopupBox";
import PageLoader from "../../common/component/PageLoader";
import { getConfidDataByID, getConfig, SaveconfigMaster, UpdateconfigMaster } from "../../../../Services/ConfigService";
import ReactTableComponent from '../ResuableComponents/ReusableDataTable';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Select from "react-select";
import { Field } from "@fluentui/react-components";
interface IConfigMaster {
    context: WebPartContext;
}

interface ILabel {
    FieldName?: string;
    ColumnType?: string;
    ListName?: string;
    DisplayColumn?: string;
    IsStaticValue?: string;
    IsShowasFilter?: string;
    Add?: string;
    AddNewRecords?: string;
    EditNewRecords?: string;
    Submit?: string;
    Update?: string;
    Cancel?: string;
    Selectanoption?: string;
    ThisFieldisRequired?: string;
    ValueAlreadyExist?: string;
    ColumnNameIsAlreadyExist?: string;
    SpecialCharacterNotAllowed?: string;
    SubmitMsg?: string;
    UpdateAlertMsg?: string;
    Atleasttwooptionrecordrequired?: string;
}

export default function ConfigMaster({ context }: IConfigMaster): JSX.Element {

    const [DisplayLabel, setDisplayLabel] = useState<ILabel>();
    const [isPanelOpen, setIsPanelOpen] = useState(false);
    const [isEditMode, setIsEditMode] = useState(false);
    const [FieldName, setFieldName] = useState("");
    const [ColumnTypeID, setColumnTypeID] = useState<any>(null);
    const [ListNameID, setListNameID] = useState('');
    const [DisplayColumnID, setDisplayColumnID] = useState<any>(null);
    const [IsShowasFilter, setIsShowasFilter] = React.useState<boolean>(false);
    const [IsStaticValue, setIsStaticValues] = React.useState<boolean>(false);
    const [options, setOptions] = React.useState<string[]>([]);
    const [newOption, setNewOption] = React.useState<string>('');
    const [ListData, setListData] = useState([]);
    const [DisplaycolumnListData, setDisplaycolumnListData] = useState([]);
    const [isToggleDisabled, setIsToggleDisabled] = useState(false);

    const [isToggleVisible, setToggleVisible] = React.useState<boolean>(false);
    const [isToggleVisible1, setToggleVisible1] = React.useState<boolean>(false);
    const [isDropdownVisible, setDropdownVisible] = React.useState<boolean>(false);
    const [isSecondaryDropdownVisible, setSecondaryDropdownVisible] = React.useState<boolean>(false);
    const [isTableVisible, setTableVisible] = React.useState<boolean>(false);
    const [showLoader, setShowLoader] = useState({ display: "none" });
    const [isPopupVisible, setisPopupVisible] = useState(false);
    const [MainTableSetdata, setData] = useState<any[]>([]);
    const [CurrentEditID, setCurrentEditID] = useState<number>(0);
    const [FieldNameErr, setFieldNameErr] = useState("");
    const [ColumnTypeIDErr, setColumnTypeIDErr] = useState("");
    const [ListNameIDErr, setListNameIDErr] = useState("");
    const [DisplayColumnIDErr, setDisplayColumnIDErr] = useState("");
    const [alertMsg, setAlertMsg] = useState("");
    const [selectedListOption, setSelectedListOption] = React.useState<any>(null);
    const inputRefs = useRef<{ [key: string]: HTMLInputElement | null; }>({});
    const [searchText, setSearchText] = useState<string>("");
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        let DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
        setDisplayLabel(DisplayLabel);
        Promise.all([getAllListFromSite(), fetchData()]).finally(() => setIsLoading(false));
    }, []);

    const fetchData = async () => {
        let FetchallConfigData: any = await getConfig(context.pageContext.web.absoluteUrl, context.spHttpClient,);
        let ConfigData = FetchallConfigData.value;
        setData(ConfigData);
    };

    const Tablecolumns = [
        {
            headerName: "Sr No",
            valueGetter: "node.rowIndex + 1",
            width: 90
        },
        {
            headerName: DisplayLabel?.FieldName || "Field Name",
            field: "Title"
        },
        {
            headerName: DisplayLabel?.ColumnType || "Column Type",
            field: "ColumnType"
        },
        {
            headerName: DisplayLabel?.ListName || "List Name",
            field: "InternalListName"
        },
        {
            headerName: DisplayLabel?.IsStaticValue || "Is Static Value",
            field: "IsStaticValue",
            cellRenderer: (params: any) => (params.value ? "Yes" : "No")
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
                    onClick={() => openEditPanel(params.data.Id)}
                />
            ),
            width: 120
        }
    ];

    const openEditPanel = async (rowData: any) => {

        setIsEditMode(true);
        setIsPanelOpen(true);

        let GetEditData = await getConfidDataByID(context.pageContext.web.absoluteUrl, context.spHttpClient, rowData);
        const EditConfigData = GetEditData.value;
        const CurrentItemId: number = EditConfigData[0].ID;
        setCurrentEditID(CurrentItemId);

        if (EditConfigData[0].ColumnType === "Dropdown" || EditConfigData[0].ColumnType === "Multiple Select")
            bindDisplayColumn(EditConfigData[0].InternalListName);

        await setFieldName(EditConfigData[0].Title);

        setColumnTypeID({ value: EditConfigData[0].ColumnType, label: EditConfigData[0].ColumnType });
        setListNameID(EditConfigData[0].InternalListName);
        setSelectedListOption({ value: EditConfigData[0].InternalListName, label: EditConfigData[0].InternalListName });
        setDisplayColumnID({ value: EditConfigData[0].DisplayValue, label: EditConfigData[0].DisplayValue });

        const TableData = (EditConfigData[0].StaticDataObject === null ? [] : EditConfigData[0].StaticDataObject.split(';'));
        await setOptions(TableData);

        if (EditConfigData[0].ColumnType === "Single line of Text") {
            setToggleVisible(false); setToggleVisible1(false);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(false); setIsToggleDisabled(false);
        } else if (EditConfigData[0].ColumnType === "Multiple lines of Text") {
            setToggleVisible(false); setToggleVisible1(false);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(false); setIsToggleDisabled(false);
        } else if (EditConfigData[0].ColumnType === "Dropdown") {
            setToggleVisible(true); setToggleVisible1(true);
            setDropdownVisible(true); setSecondaryDropdownVisible(true);
            setTableVisible(false); setIsToggleDisabled(false);
        } else if (EditConfigData[0].ColumnType === "Multiple Select") {
            setToggleVisible(true); setToggleVisible1(true);
            setDropdownVisible(true); setSecondaryDropdownVisible(true);
            setTableVisible(false); setIsToggleDisabled(false);
        } else if (EditConfigData[0].ColumnType === "Radio") {
            setToggleVisible(true); setToggleVisible1(true);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(true); setIsStaticValues(true); setIsToggleDisabled(true);
        } else if (EditConfigData[0].ColumnType === "Date and Time") {
            setToggleVisible(true); setToggleVisible1(false);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(false); setIsToggleDisabled(false);
        } else if (EditConfigData[0].ColumnType === "Person or Group") {
            setToggleVisible(true); setToggleVisible1(false);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(false); setIsToggleDisabled(false);
        } else {
            setToggleVisible(false); setToggleVisible1(false);
            setDropdownVisible(false); setSecondaryDropdownVisible(false);
            setTableVisible(false); setIsToggleDisabled(false);
        }

        await setIsShowasFilter(EditConfigData[0].IsShowAsFilter);
        await setIsStaticValues(EditConfigData[0].IsStaticValue);

        if (EditConfigData[0].IsStaticValue === true) {
            setDropdownVisible(false);
            setSecondaryDropdownVisible(false);
            setTableVisible(true);
        }
    };

    const [newOptionError, setNewOptionErr] = useState("");

    const addOption = () => {
        setNewOptionErr("");
        if (newOption.trim() === '') {
            setNewOptionErr(DisplayLabel?.ThisFieldisRequired || "");
            return;
        }
        const isDuplicate = options.some(
            (Data) => Data.toLowerCase() === newOption.toLowerCase().trim()
        );
        if (isDuplicate) {
            setNewOptionErr(DisplayLabel?.ValueAlreadyExist || "");
            return;
        }
        setOptions([...options, newOption.trim()]);
        setNewOption('');
    };

    const removeOption = (index: number) => {
        setOptions(options.filter((_, i) => i !== index));
    };

    async function getAllListFromSite() {
        var url = context.pageContext.web.absoluteUrl + "/_api/web/lists?$select=Title&$filter=(Hidden eq false) and (BaseType ne 1) and Title ne 'ConfigEntryMaster'";
        const data = await GetListData(url);
        var ListNamedata = data.d.results;
        let options = ListNamedata.map((item: any) => ({ value: item.Title, label: item.Title }));
        setListData(options);
    }

    async function GetListData(query: string) {
        const response = await context.spHttpClient.get(query, SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'odata-version': '',
            },
        });
        return await response.json();
    }

    const openAddPanel = () => {
        clearField();
        setIsEditMode(false);
        setIsPanelOpen(true);
    };

    const closePanel = () => {
        setIsPanelOpen(false);
    };

    const handleIsShowasFilterToggleChange = (checked: boolean): void => {
        setIsShowasFilter(checked);
    };

    const handleIsStaticValueToggleChange = (checked: boolean): void => {
        setIsStaticValues(checked);
        if (checked) {
            setDropdownVisible(false);
            setSecondaryDropdownVisible(false);
            setTableVisible(true);
        } else {
            setDropdownVisible(true);
            setSecondaryDropdownVisible(true);
            setTableVisible(false);
        }
    };

    const dropdownOptions = [
        { value: 'Single line of Text', label: 'Single line of Text' },
        { value: 'Multiple lines of Text', label: 'Multiple lines of Text' },
        { value: 'Dropdown', label: 'Dropdown' },
        { value: 'Multiple Select', label: 'Multiple Select' },
        { value: 'Radio', label: 'Radio' },
        { value: 'Date and Time', label: 'Date and Time' },
        { value: 'Person or Group', label: 'Person or Group' },
    ];

    const handleColumnTypeonChange = (option?: any) => {
        setColumnTypeID(option);
        if (option) {
            if (option.value === "Single line of Text") {
                setToggleVisible(false); setToggleVisible1(false);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(false); setIsToggleDisabled(false);
            } else if (option.value === "Multiple lines of Text") {
                setToggleVisible(false); setToggleVisible1(false);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(false); setIsToggleDisabled(false);
            } else if (option.value === "Dropdown") {
                setToggleVisible(true); setToggleVisible1(true);
                setDropdownVisible(true); setSecondaryDropdownVisible(true);
                setTableVisible(false); setIsToggleDisabled(false);
            } else if (option.value === "Multiple Select") {
                setToggleVisible(true); setToggleVisible1(true);
                setDropdownVisible(true); setSecondaryDropdownVisible(true);
                setTableVisible(false); setIsToggleDisabled(false);
            } else if (option.value === "Radio") {
                setToggleVisible(true); setToggleVisible1(true);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(true); setIsStaticValues(true); setIsToggleDisabled(true);
            } else if (option.value === "Date and Time") {
                setToggleVisible(true); setToggleVisible1(false);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(false); setIsToggleDisabled(false);
            } else if (option.value === "Person or Group") {
                setToggleVisible(true); setToggleVisible1(false);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(false); setIsToggleDisabled(false);
            } else {
                setToggleVisible(false); setToggleVisible1(false);
                setDropdownVisible(false); setSecondaryDropdownVisible(false);
                setTableVisible(false); setIsToggleDisabled(false);
            }
        }
    };

    const handleListNameonChange = async (option?: any) => {
        bindDisplayColumn(option.label);
        setListNameID(option.value);
        setSelectedListOption(option);
    };

    const bindDisplayColumn = async (listName: string) => {
        let query = context.pageContext.web.absoluteUrl + "/_api/web/lists/getbytitle('" + listName + "')/Fields?$filter=(CanBeDeleted eq true) and (TypeAsString eq 'Text' or TypeAsString eq 'Number')";
        const data = await GetListData(query);
        let DisplayColumnData = data.d.results;
        const optionsData: any = DisplayColumnData.map((item: any) => ({ value: item.Title, label: item.Title }));
        setDisplaycolumnListData(optionsData);
    };

    const handleDisplayColumnonChange = async (option?: any) => {
        setDisplayColumnID(option);
    };

    const hidePopup = React.useCallback(() => {
        setisPopupVisible(false);
        clearField();
        closePanel();
        setShowLoader({ display: "none" });
    }, [isPopupVisible]);

    const clearField = () => {
        setCurrentEditID(0);
        setFieldName("");
        setColumnTypeID(null);
        setListNameID('');
        setSelectedListOption(null);
        setDisplayColumnID(null);
        setIsShowasFilter(false);
        setIsStaticValues(false);
        setOptions([]);
        clearError();
        setToggleVisible(false);
        setToggleVisible1(false);
        setDropdownVisible(false);
        setSecondaryDropdownVisible(false);
        setTableVisible(false);
        setIsToggleDisabled(false);
    };

    const clearError = () => {
        setFieldNameErr("");
        setColumnTypeIDErr("");
        setListNameIDErr("");
        setDisplayColumnIDErr("");
    };

    const validation = () => {
        let isValidForm = true;

        if (FieldName === "" || FieldName === undefined || FieldName === null) {
            setFieldNameErr(DisplayLabel?.ThisFieldisRequired as string);
            inputRefs.current["FieldName"]?.focus();
            isValidForm = false;
            return;
        }
        const isDuplicate = MainTableSetdata.some(
            (Data) => Data.Title.toLowerCase() === FieldName.toLowerCase().trim()
        );
        if (/[*|\":<>[\]{}`\\()'!%;@#&$]/.test(FieldName)) {
            setFieldNameErr(DisplayLabel?.SpecialCharacterNotAllowed as string);
            isValidForm = false;
            return;
        }
        if (isDuplicate && !isEditMode) {
            setFieldNameErr(DisplayLabel?.ColumnNameIsAlreadyExist as string);
            isValidForm = false;
            return;
        }
        if (isDuplicate && isEditMode) {
            MainTableSetdata.map((Data) => {
                if (Data.Title.toLowerCase() === FieldName.toLowerCase().trim() && Data.ID !== CurrentEditID) {
                    setFieldNameErr(DisplayLabel?.ColumnNameIsAlreadyExist as string);
                    isValidForm = false;
                    return;
                }
            });
        }
        if (ColumnTypeID?.value === "" || ColumnTypeID?.value === undefined || ColumnTypeID === null) {
            setColumnTypeIDErr(DisplayLabel?.ThisFieldisRequired as string);
            inputRefs.current["ColumnType"]?.focus();
            isValidForm = false;
            return;
        }
        if (IsStaticValue === true) {
            if (options.length === 0) {
                alert(DisplayLabel?.Atleasttwooptionrecordrequired);
                isValidForm = false;
                return;
            }
        }
        if (IsStaticValue === false && ColumnTypeID.value === "Dropdown") {
            if (ListNameID === "" || ListNameID === undefined || ListNameID === null) {
                setListNameIDErr(DisplayLabel?.ThisFieldisRequired as string);
                inputRefs.current["ListName"]?.focus();
                isValidForm = false;
                return;
            }
            if (DisplayColumnID === "" || DisplayColumnID === undefined || DisplayColumnID === null) {
                setDisplayColumnIDErr(DisplayLabel?.ThisFieldisRequired as string);
                inputRefs.current["DisplayColumn"]?.focus();
                isValidForm = false;
                return;
            }
        }

        return isValidForm;
    };

    const SaveItemData = () => {
        clearError();
        let valid = validation();
        valid ? saveData() : "";
    };

    const saveData = async () => {
        try {
            setShowLoader({ display: "block" });

            let ddlListName = null;
            let ddlColumn = null;

            if (IsStaticValue === true) {
                ddlListName = null;
                ddlColumn = null;
            } else {
                ddlListName = ListNameID;
                ddlColumn = DisplayColumnID?.value || "";
            }

            let FieldNameNew = FieldName.split(" ").join("");
            let Name = FieldName;

            let option = {
                __metadata: { type: "SP.Data.ConfigEntryMasterListItem" },
                Title: Name.trim(),
                InternalTitleName: FieldNameNew,
                IsActive: true,
                ColumnType: ColumnTypeID.value,
                IsStaticValue: IsStaticValue,
                StaticDataObject: options.join(';'),
                InternalListName: ddlListName,
                DisplayValue: ddlColumn,
                IsShowAsFilter: IsShowasFilter,
                Abbreviation: "Abbreviation"
            };

            if (!isEditMode)
                await SaveconfigMaster(context.pageContext.web.absoluteUrl, context.spHttpClient, option);
            else
                await UpdateconfigMaster(context.pageContext.web.absoluteUrl, context.spHttpClient, option, CurrentEditID);

            setShowLoader({ display: "none" });
            setIsPanelOpen(false);
            setAlertMsg((isEditMode ? DisplayLabel?.UpdateAlertMsg : DisplayLabel?.SubmitMsg) || "");
            setisPopupVisible(true);
            fetchData();

        } catch (error) {
            console.error("Error during save operation:", error);
            setShowLoader({ display: "none" });
        }
    };

    const FilterMainTableSetdata = MainTableSetdata.filter((items) => {
        const terms = searchText.toLowerCase().split(' ').filter(Boolean);

        const date = items.Created ? new Date(items.Created) : null;
        let formattedDate = "";
        if (date) {
            const day = ("0" + date.getDate()).slice(-2);
            const month = ("0" + (date.getMonth() + 1)).slice(-2);
            const year = date.getFullYear();
            let hours = date.getHours();
            const minutes = ("0" + date.getMinutes()).slice(-2);
            const ampm = hours >= 12 ? "PM" : "AM";
            hours = hours % 12;
            hours = hours ? hours : 12;
            formattedDate = `${day}/${month}/${year} at ${hours}:${minutes} ${ampm}`;
        }

        const searchableString = [
            items.ColumnType,
            items.InternalTitleName,
            items.InternalListName,
            items.IsStaticValue ? "Yes" : "No",
            formattedDate,
            formattedDate.replace(/\//g, "-"),
            formattedDate.replace(" at ", " "),
        ]
            .map(val => (val ? String(val).toLowerCase() : ''))
            .join(' ');

        return terms.every(term => searchableString.includes(String(term).toLowerCase()));
    });

    if (isLoading) {
        return <PageLoader message="Loading configuration..." minHeight="72vh" />;
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

                    <PrimaryButton text={DisplayLabel?.Add || "Add Config"} onClick={openAddPanel} />

                </div>

                <ReactTableComponent
                    rowData={FilterMainTableSetdata}
                    columnDefs={Tablecolumns}
                />

            </Stack>

            <Panel
                isOpen={isPanelOpen}
                onDismiss={closePanel}
                closeButtonAriaLabel="Close"
                type={PanelType.large}
                isFooterAtBottom
                headerText={isEditMode ? DisplayLabel?.EditNewRecords || "Edit Config" : DisplayLabel?.AddNewRecords || "Add Config"}
                onRenderFooterContent={() => (
                    <>
                        <PrimaryButton
                            text={isEditMode ? DisplayLabel?.Update || "Update" : DisplayLabel?.Submit || "Submit"}
                            onClick={SaveItemData}
                            styles={{ root: { marginRight: 8 } }}
                        />
                        <DefaultButton text={DisplayLabel?.Cancel || "Cancel"} onClick={closePanel} />
                    </>
                )}
            >

                <div className="grid-2">
                    <Field>
                        <label>{DisplayLabel?.FieldName || "Field Name"} <span style={{ color: "red" }}>*</span></label>
                        <TextField
                            placeholder="Enter Field Name"
                            value={FieldName}
                            onChange={(_, value) => setFieldName(value || "")}
                            componentRef={(input: any) => (inputRefs.current["FieldName"] = input)}
                        />
                        {FieldNameErr && (
                            <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>{FieldNameErr}</p>
                        )}
                    </Field>
                    <Field>
                        <label className="Headerlabel">
                            {DisplayLabel?.ColumnType || "Column Type"} <span style={{ color: "red" }}>*</span>
                        </label>
                        <Select
                            options={dropdownOptions}
                            value={ColumnTypeID}
                            onChange={(selected: any) => handleColumnTypeonChange(selected)}
                            placeholder="Select Column Type"
                        />
                        {ColumnTypeIDErr && (
                            <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>{ColumnTypeIDErr}</p>
                        )}
                    </Field>
                </div>
                {isToggleVisible && isToggleVisible1 && (
                    <div className="grid-2" >

                        <Field>
                            <label className="Headerlabel">
                                {DisplayLabel?.IsShowasFilter || "Is Show as Filter"}
                            </label>
                            <Toggle
                                checked={IsShowasFilter}
                                onChange={(_, checked) => handleIsShowasFilterToggleChange(checked!)}
                            />
                        </Field>

                        <Field>
                            <label className="Headerlabel">
                                {DisplayLabel?.IsStaticValue || "Is Static Value"}
                            </label>
                            <Toggle
                                checked={IsStaticValue}
                                onChange={(_, checked) => handleIsStaticValueToggleChange(checked!)}
                                disabled={isToggleDisabled}
                            />
                        </Field>

                    </div>
                )}



                {isDropdownVisible && isSecondaryDropdownVisible && (
                    <div className="grid-2">

                        <Field>
                            <label className="Headerlabel">
                                {DisplayLabel?.ListName || "List Name"} <span style={{ color: "red" }}>*</span>
                            </label>
                            <Select
                                options={ListData}
                                value={selectedListOption}
                                onChange={(selected: any) => handleListNameonChange(selected)}
                                placeholder="Select List"
                            />
                            {ListNameIDErr && <p style={{ color: "red", fontSize: 12 }}>{ListNameIDErr}</p>}
                        </Field>

                        <Field>
                            <label className="Headerlabel">
                                {DisplayLabel?.DisplayColumn || "Display Column"} <span style={{ color: "red" }}>*</span>
                            </label>
                            <Select
                                options={DisplaycolumnListData}
                                value={DisplayColumnID}
                                onChange={(selected: any) => handleDisplayColumnonChange(selected)}
                                placeholder="Select Column"
                            />
                            {DisplayColumnIDErr && <p style={{ color: "red", fontSize: 12 }}>{DisplayColumnIDErr}</p>}
                        </Field>

                    </div>
                )}


                {isTableVisible && (
                    <div>
                        <table className="fluent-table">
                            <thead>
                                <tr>
                                    <th style={{ textAlign: "left", width: "10%" }}>Sr. No.</th>
                                    <th style={{ textAlign: "left", width: "70%" }}>Option <span style={{ color: "red" }}>*</span></th>
                                    <th style={{ textAlign: "center", width: "20%" }}>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td />
                                    <td style={{ paddingTop: 8 }}>
                                        <TextField
                                            placeholder="Enter Option"
                                            value={newOption}
                                            onChange={(_, value) => setNewOption(value || '')}
                                            styles={{ root: { width: "100%" } }}
                                            errorMessage={newOptionError}
                                        />
                                    </td>
                                    <td style={{ textAlign: "center" }}>
                                        <FontIcon
                                            iconName="Add"
                                            style={{
                                                color: "#fff",
                                                cursor: "pointer",
                                                backgroundColor: "#009ef7",
                                                padding: "4px 8px",
                                                borderRadius: "50%"
                                            }}
                                            onClick={addOption}
                                        />
                                    </td>
                                </tr>
                            </tbody>
                            <tbody>
                                {options.map((option, index) => (
                                    <tr key={index} style={{ borderBottom: "1px solid #ddd" }}>
                                        <td style={{ padding: 8 }}>{index + 1}</td>
                                        <td style={{ padding: 8 }}>{option}</td>
                                        <td style={{ textAlign: "center" }}>
                                            <FontIcon
                                                iconName="Delete"
                                                style={{
                                                    color: "#f1416c",
                                                    cursor: "pointer",
                                                    backgroundColor: "#f5f8fa",
                                                    padding: "6px 9px",
                                                    borderRadius: "4px"
                                                }}
                                                onClick={() => removeOption(index)}
                                            />
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}



            </Panel>

            <PopupBox isPopupBoxVisible={isPopupVisible} hidePopup={hidePopup} msg={alertMsg} type={isEditMode ? "update" : "insert"} />

        </div>
    );
}
