import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useEffect, useState } from 'react';
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import {
    Checkbox,
    ChoiceGroup,
    DefaultButton,
    Dropdown,
    PrimaryButton,
    TextField,
    Toggle,
} from '@fluentui/react';
import {
    ChevronUp20Regular,
    ChevronDown20Regular,
    Add20Regular,
    Edit16Regular,
    Delete16Regular,
    TabDesktop20Regular,
} from '@fluentui/react-icons';
import { ILabel } from "../../../../Intrface/ILabel";
import Select, { CSSObjectWithLabel } from "react-select";
import { getConfigActive } from "../../../../Services/ConfigService";
import { getRoles } from "../../../../Services/Role";
import { getAllButtons } from "../../../../Services/Buttons";
import { IButtonsProps, IRolePermission } from "../../../../Intrface/IButtonInterface";
import { format } from "date-fns";
import { getActiveRedundancyDays } from "../../../../Services/ArchiveRedundancyDaysService";
import { getUserIdFromLoginName, uuidv4 } from "../../../../DAL/Commonfile";
import { getListData } from "../../../../Services/GeneralDocument";
import { getDataById, getTileAllData, SaveTileSetting, UpdateTileSetting } from "../../../../Services/MasTileService";
import { createColumn, getColumnType, GetListData, TileLibrary } from "../../common/ListCreation";
import { breakRoleInheritanceForLib, grantPermissionsForLib } from "../../../../Services/FolderStructure";
import { spfi, SPFx } from '@pnp/sp';
import { IColumnSchema } from "../../../../Intrface/IListSchema";
import PageLoader from "../../common/component/PageLoader";
import FieldError from "../../common/component/FieldError";
import { getPrimaryActionButtonStyles, getSecondaryActionButtonStyles } from "../../common/component/buttonStyles";

interface ITileFormProps {
    context: WebPartContext;
    tileID: number;
    setIsOpenEditor: (value: boolean) => void;
    allTiles: any[];
}

const TileForm: React.FunctionComponent<ITileFormProps> = ({ context, setIsOpenEditor, tileID, allTiles }) => {
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const SiteURL = context.pageContext.web.absoluteUrl;
    const refrenceNOData = `${format(new Date(), 'yyyy')}-00001`;
    const [allButtonsWithPermissions, setAllButtonsWithPermissions] = useState<IRolePermission[]>([]);
    const [fieldData, setFieldData] = useState<any[]>([]);
    const [isEditMode, setIsEditMode] = useState(false);
    const [formData, setFormData] = useState<Record<string, any>>({
        field: null,
        IsRequired: false,
        IsActiveControl: true,
        IsFieldAllowInFile: false,
        isShowAsFilter: false,
        Flag: "New",
        editingIndex: -1,
        isTileStatus: true,
        increment: "Continue",
        separator: "-",
        Archive: "Archive",
        isAllowApprover: false,
        allowChildInheritance: false,
        isArchiveAllowed: false,
        isDynamicReference: false
    });
    const [refFormatData, setRefFormatData] = useState<string[]>([]);
    const [errors, setErrors] = useState<Record<string, any>>({});
    const [refExample, setRefExample] = useState<string>(refrenceNOData);
    const [roles, setRoles] = useState<any[]>([]);
    const [allButtons, setAllButtons] = useState<IButtonsProps[]>([]);
    const [groupByButtons, setGroupByButtons] = useState<any[]>([]);
    const [customSeparators, setCustomSeparators] = useState<{ [key: number]: string; }>({});
    const [redundancyData, setRedundancyData] = useState([]);
    const [tableData, setTableData] = useState<any[]>([]);
    const [admin, setAdmin] = useState<any[]>([]);
    const [isDisabled, setIsDisabled] = useState<boolean>(false);
    const [expandedSections, setExpandedSections] = useState<Set<string>>(
        new Set(['tileDetails', 'fields', 'referenceNo', 'archive', 'buttons'])
    );
    const [isToggleDisabled, setIsToggleDisabled] = useState(false);
    const [isPageLoading, setIsPageLoading] = useState(true);

    const RedundancyDaysData = async () => {
        const ActiveRedundancyDaysData: any = await getActiveRedundancyDays(SiteURL, context.spHttpClient);
        const options: any = ActiveRedundancyDaysData.value.map((item: any) => ({ value: item.ID, label: item.RedundancyDays }));
        setRedundancyData(options);
        return options;
    };


    const toggleSection = (key: string) => {
        setExpandedSections(prev => {
            const newSet = new Set(prev);
            if (newSet.has(key)) newSet.delete(key);
            else newSet.add(key);
            return newSet;
        });
    };

    useEffect(() => {
        const initializeForm = async () => {
            setIsPageLoading(true);
            const redundancyOptions = await RedundancyDaysData();
            await Promise.all([
                fetchFieldConfig(),
                getAllRoles(),
                getAllButton(),
                getAdmin()
            ]);
            if (tileID > 0) {
                await openEditPanel(redundancyOptions);
            }
            setIsPageLoading(false);
        };

        void initializeForm();
    }, [tileID]);

    const openEditPanel = async (redundancyOptionsParam?: any[]) => {

        const GetEditData = await getDataById(SiteURL, context.spHttpClient, tileID);
        const EditSettingData = GetEditData.value[0];

        const TileAdminData: any = EditSettingData?.TileAdmin ? ([EditSettingData?.TileAdmin.EMail]) : [];

        getAllColumns(EditSettingData?.LibraryName);

        setTableData(EditSettingData?.DynamicControl === null ? [] : JSON.parse(EditSettingData?.DynamicControl));

        if (EditSettingData?.IsArchiveRequired === true) {
            const currentRedundancyData = redundancyOptionsParam || redundancyData;
            const FilterRetentionDays = currentRedundancyData.find((item: any) => item.label === EditSettingData?.RetentionDays);
            setFormData((prevData) => ({
                ...prevData, RedundancyData: FilterRetentionDays,
                ArchiveInternal: EditSettingData?.ArchiveLibraryName,
                ArchiveVersions: EditSettingData?.ArchiveVersionCount
            }));
        }

        if (EditSettingData.IsDynamicReference) {
            const formula = EditSettingData.ReferenceFormula || "";
            const fields = new Set<string>();
            const extracted = formula.match(/\{[^}]+\}/g);

            extracted.map((m: any, index: number) => {
                const fieldName = m.replace(/[{}]/g, "");
                if (index === extracted.length - 1)
                    setFormData((prevData) => ({ ...prevData, increment: fieldName || "Continue" }));
                else
                    fields.add(fieldName);

            }) || [];
            setRefExample(formula);
            setRefFormatData(Array.from(fields));
        }
        else {
            setRefExample(EditSettingData.ReferenceFormula);
        }

        setIsEditMode(true);

        const permissionData = EditSettingData?.Permission || [];
        const PermissionIds: number[] = permissionData.map((person: any) => person.Id);

        const PermissionEmails = permissionData.map((p: any) => {
            if (!p.Name) return "";

            if (p.Name.includes("membership")) {
                return p.Name.split('|').pop();
            }

            return p.Title;
        });
        setAllButtonsWithPermissions(EditSettingData?.CustomPermission ? JSON.parse(EditSettingData?.CustomPermission) : []);
        setFormData((prevData) => ({
            ...prevData,
            TileName: EditSettingData?.TileName,
            TileAdminId: EditSettingData?.TileAdmin?.Id,
            TileAdminEmail: TileAdminData,
            isTileStatus: EditSettingData?.Active,
            isAllowApprover: EditSettingData?.AllowApprover,
            allowChildInheritance: EditSettingData?.AllowChildInheritance,
            isArchiveAllowed: EditSettingData?.IsArchiveRequired,
            isDynamicReference: EditSettingData?.IsDynamicReference,
            separator: EditSettingData?.Separator || "-",
            PermissionEmail: PermissionEmails,
            PermissionIds: PermissionIds
        }));
    };


    const getAdmin = async () => {
        const data = await getListData(`${SiteURL}/_api/web/lists/getbytitle('DMS_GroupName')/items`, context);
        setAdmin(data.value.map((el: any) => (el.GroupNameId)));
    };
    const fetchFieldConfig = async () => {
        const ConfigData: any = await getConfigActive(SiteURL, context.spHttpClient);
        const options: any = ConfigData.value.map((item: any) => ({ value: item.ID, label: item.Title, ...item }));
        setFieldData(options);
    };

    const getAllRoles = () => {
        return getRoles(SiteURL, context.spHttpClient).then((response: any) => {
            const roleData = response.value;
            setRoles(roleData);
        });
    };
    const getAllButton = () => {
        return getAllButtons(SiteURL, context.spHttpClient).then((response: any) => {
            const buttonData = response.value;
            const buttonsList: IButtonsProps[] = buttonData.map(convertToButtonProps);
            const grouped = buttonsList.reduce((acc: any, item: any) => {
                (acc[item.ButtonType] = acc[item.ButtonType] || []).push(item);
                return acc;
            }, {} as { [key: string]: any[]; });

            setGroupByButtons(grouped);
            setAllButtons(buttonsList);
        });
    };

    const convertToButtonProps = (backendData: any) => (
        {
            Id: backendData.Id,
            ButtonDisplayName: backendData.ButtonDisplayName,
            ButtonType: backendData.ButtonType,
            Title: backendData.Title,
            key: backendData.InternalName,
            value: false,
            Active: backendData.Active,
            InternalName: backendData.InternalName,
        }
    );
    useEffect(() => {
        if (allButtons.length > 0 && roles.length > 0) {
            bindAllButtons();
        }
    }, [allButtons, roles]);

    const bindAllButtons = () => {
        let buttonObj: IRolePermission = { Role: "", Permission: [], UsersId: [] };
        const roleD = roles.map((role: any) => {

            buttonObj = { Role: role.Title, Permission: allButtons, UsersId: [] };
            return buttonObj;
        });
        setAllButtonsWithPermissions(roleD);
    };


    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any
    };


    const handleFieldChange = async (option: any) => {
        setFormData((prevData) => ({
            ...prevData, field: option.value,
            IsRequired: false,
            IsActiveControl: true,
            IsFieldAllowInFile: false,
            isShowAsFilter: false,
        }));

        const selectedOption = fieldData.find((element: any) => element.ID === option.value);

        if (selectedOption) {
            if (selectedOption.IsShowAsFilter) {
                setIsToggleDisabled(false);
            } else {
                setIsToggleDisabled(true);
            }
        }
    };

    const handleInputChange = (key: string, value: any) => {
        setFormData({ ...formData, [key]: value });
    };

    const handleSave = async () => {
        setErrors((prevData) => ({ ...prevData, field: '' }));
        if (formData.editingIndex >= 0) {
            const updatedData = [...tableData];
            updatedData[formData.editingIndex] = { ...formData };
            delete updatedData[formData.editingIndex].editingIndex;
            setTableData(updatedData);

        } else {
            if (formData.field !== null) {
                const selectedOption: any = fieldData.find((element: any) => element.ID === formData.field);

                const isDuplicate = tableData.find((element: any) => element.field === formData.field);

                if (isDuplicate === undefined) {

                    setTableData((prevData: any[]) => [
                        ...prevData,
                        { ...formData, ...selectedOption },
                    ]);
                }

                else {
                    setErrors((prevData) => ({ ...prevData, field: 'duplicate' }));
                }
            }
            else {
                // alert("Please Select");
                setErrors((prevData) => ({ ...prevData, field: 'Please Select' }));
            }
        }

        setFormData((prevData) => ({
            ...prevData,
            field: null,
            IsRequired: false,
            IsActiveControl: true,
            IsFieldAllowInFile: false,
            isShowAsFilter: false,
            Flag: "New",
            editingIndex: -1,
        }));
    };

    const handleEdit = (index: number) => {
        setFormData({ ...tableData[index], editingIndex: index });
    };

    const handleDelete = (index: number) => {
        const updatedData = tableData.filter((_, i) => i !== index);
        setTableData(updatedData);
    };
    const handleCheckboxToggle = (item: string, isChecked: boolean) => {
        const updatedRefData = isChecked
            ? [...refFormatData, item]
            : refFormatData.filter((refItem) => refItem !== item);
        setRefFormatData(updatedRefData);
        generateFormula(updatedRefData, formData?.prefix, formData?.separator, formData?.increment);
    };

    const grantButtonPermission = (role: string, permission: any) => {
        setAllButtonsWithPermissions((prev: IRolePermission[]) =>
            prev.map((r) =>
                r.Role === role
                    ? {
                        ...r,
                        UsersId: permission,
                    }
                    : r
            )
        );
    };

    const handleCheckboxChangeButton = (role: string, key: Number, val: boolean) => {
        setAllButtonsWithPermissions((prev: IRolePermission[]) =>
            prev.map((r) =>
                r.Role === role
                    ? {
                        ...r,
                        Permission: r.Permission.map((p: any) =>
                            p.Id === key ? { ...p, value: val } : p
                        ),
                    }
                    : r
            )
        );
    };

    const bindPermission = (role: string) => {
        const foundRole = allButtonsWithPermissions.find((r) => r.Role === role);
        const userEMail = foundRole ? foundRole.UsersId.map((user: any) => (user.secondaryText || user.loginName || user.email)) : [];
        return userEMail;
    };

    const generateFormula = (
        refData: string[],
        prefixValue: string,
        separatorValue: string,
        incrementValue: string,
        customSeparatorData: { [key: number]: string; } = customSeparators
    ) => {
        let formula = prefixValue ? `${prefixValue}${separatorValue}` : "";

        refData.forEach((item, index) => {
            formula += `{${item}}`;

            if ((customSeparatorData[index] || "Separator") === "Separator") {
                formula += separatorValue;
            }
        });

        if (formula.endsWith(separatorValue)) {
            formula = formula.slice(0, -separatorValue.length);
        }

        formula += separatorValue;
        formula += `{${incrementValue}}`;
        setRefExample(formula);
    };

    const handleRadioChange = (type: string, value: string) => {
        if (type === "separator") {
            setFormData(prevData => ({ ...prevData, separator: value }));
            generateFormula(refFormatData, formData?.prefix, value, formData?.increment);
        } else if (type === "increment") {
            setFormData(prevData => ({ ...prevData, increment: value }));
            generateFormula(refFormatData, formData?.prefix, formData?.separator, value);
        }
    };
    const handlePrefixChange = (value: string) => {
        setFormData(prevData => ({ ...prevData, prefix: value }));
        generateFormula(refFormatData, value, formData?.separator, formData?.increment);
    };


    const CheckboxData = (obj: any) => {
        let icheckbox;
        if (obj.ColumnType === 'Dropdown' && !obj.IsStaticValue && obj.IsRequired === true && obj.IsFieldAllowInFile != true && obj.IsActiveControl === true) {
            icheckbox = <Checkbox label={obj.Title} checked={refFormatData.includes(obj.Title)} onChange={(e, checked) => handleCheckboxToggle(obj.Title, checked!)} />;
        }
        return icheckbox;
    };

    const renderFieldTable = () => (
        <table className="fluent-table">
            <thead>
                <tr >
                    <th>{DisplayLabel.SrNo}</th>
                    <th>{DisplayLabel.Field}</th>
                    <th>{DisplayLabel.IsRequired}</th>
                    <th>{DisplayLabel.FieldStatus}</th>
                    <th>{DisplayLabel.IsFieldAllowinFile}</th>
                    <th>{DisplayLabel.SearchFilterRequired}</th>
                    <th>{DisplayLabel.Action}</th>
                </tr>


                <tr>
                    <td></td>
                    <td>
                        <Select
                            options={fieldData}
                            value={fieldData.find((option: any) => option.value === formData.field) || {}}
                            onChange={handleFieldChange}
                            isSearchable
                            placeholder={DisplayLabel?.Selectanoption}
                        />
                        <span style={{ color: "#a4262c" }}>{errors.field}</span>
                    </td>
                    <td><Toggle className="tile-field-toggle" checked={formData.IsRequired} onChange={(e, checked) => handleInputChange('IsRequired', checked)} /></td>
                    <td><Toggle className="tile-field-toggle" checked={formData.IsActiveControl} onChange={(e, checked) => handleInputChange('IsActiveControl', checked)} /></td>
                    <td><Toggle className="tile-field-toggle" checked={formData.IsFieldAllowInFile} onChange={(e, checked) => handleInputChange('IsFieldAllowInFile', checked)} /></td>
                    <td><Toggle className="tile-field-toggle" checked={formData.isShowAsFilter} onChange={(e, checked) => handleInputChange('isShowAsFilter', checked)} disabled={isToggleDisabled} /></td>
                    <td className="tile-field-col-action">
                        <button className="tile-field-add-btn">
                            <Add20Regular onClick={() => handleSave()} />
                        </button>
                    </td>
                </tr>
            </thead>
            <tbody>
                {tableData.map((row, index) => (
                    <tr key={index}>
                        <td>{index + 1}</td>
                        <td>{row.Title}</td>
                        <td>{row.IsRequired ? 'Yes' : 'No'}</td>
                        <td>{row.IsActiveControl ? 'Yes' : 'No'}</td>
                        <td>{row.IsFieldAllowInFile ? 'Yes' : 'No'}</td>
                        <td>{row.isShowAsFilter ? 'Yes' : 'No'}</td>
                        <td>
                            <span className="tile-field-action-btns">
                                <button className="tile-field-edit-btn" onClick={() => handleEdit(index)}>
                                    <Edit16Regular />
                                </button>
                                {row.Flag && (
                                    <button className="tile-field-delete-btn" onClick={() => handleDelete(index)} >
                                        <Delete16Regular />
                                    </button>
                                )}
                            </span>
                        </td>

                    </tr>
                ))}
            </tbody>
        </table>
    );

    const handleDropdownChange = (index: number, value: string) => {
        const updatedSeparators = { ...customSeparators, [index]: value };
        setCustomSeparators(updatedSeparators);
        generateFormula(refFormatData, formData?.prefix, formData?.separator, formData?.increment, updatedSeparators);
    };

    const handleArchiveDropdownChange = (option?: any) => {
        setFormData((prevData) => ({ ...prevData, RedundancyData: option }));
    };

    const renderRefForm = () => {
        return <div>
            {
                formData?.isDynamicReference && (
                    <div style={{ marginBottom: '20px' }}>

                        <label style={{ marginBottom: '10px', display: 'block' }}>{DisplayLabel?.ChooseFields}</label>
                        <div className="row">
                            <Checkbox className="column2" label="YYYY" checked={refFormatData.includes("YYYY")} onChange={(e, checked) => handleCheckboxToggle("YYYY", checked!)} />
                            <Checkbox className="column2" label="YY_YY" checked={refFormatData.includes("YY_YY")} onChange={(e, checked) => handleCheckboxToggle("YY_YY", checked!)} />
                            <Checkbox className="column2" label="MM" checked={refFormatData.includes("MM")} onChange={(e, checked) => handleCheckboxToggle("MM", checked!)} />
                            {
                                tableData.map((el) => (CheckboxData(el)))
                            }
                        </div>
                    </div>
                )
            }
            {
                formData?.isDynamicReference && (
                    <div >
                        <div
                            style={{
                                display: 'flex',
                                flexDirection: 'column',
                                gap: '10px',
                                padding: '10px',

                            }}
                        >
                            <label style={{ display: 'block' }}>{DisplayLabel?.Separator}</label>
                            <ChoiceGroup
                                options={[
                                    { key: "-", text: "Hyphens ( - )" },
                                    { key: "/", text: "Slash ( / )" },
                                ]}
                                selectedKey={formData?.separator}
                                onChange={(e, option) => {
                                    handleRadioChange("separator", option?.key!);
                                }}
                                required={true}
                                className="row"
                                styles={{
                                    flexContainer: {
                                        display: "flex",
                                        flexDirection: "row",
                                        gap: "10px",
                                        flexWrap: 'wrap'
                                        /* backgroundColor: "#f5f8fa",*/
                                    },
                                }}
                            />
                        </div>

                        {/* Initial Increment Choice Group */}
                        <div
                            style={{
                                display: 'flex',
                                flexDirection: 'column',
                                gap: '10px',
                                padding: '10px',
                                flex: 1,
                            }}
                        >
                            <label style={{ display: 'block' }}>{DisplayLabel?.InitialIncrement}</label>

                            <ChoiceGroup
                                options={[
                                    { key: "Continue", text: "Continue" },
                                    { key: "Monthly", text: "Monthly" },
                                    { key: "Yearly", text: "Yearly" },
                                    { key: "Financial Year", text: "Financial Year" },
                                    { key: "Manual", text: "Manual" },
                                ]}
                                selectedKey={formData?.increment}
                                onChange={(e, option) => {
                                    handleRadioChange("increment", option?.key!);
                                }}
                                required={true}
                                styles={{
                                    flexContainer: {
                                        display: "flex",
                                        flexDirection: "row",
                                        gap: "10px",
                                        flexWrap: 'wrap'
                                    },
                                }}
                            />
                        </div>
                    </div>
                )
            }

            {
                formData?.isDynamicReference && (
                    <div>
                        {/* Choose Fields Section */}

                        <div>
                            <label style={{ display: 'block' }}>{DisplayLabel?.ChangeSetting}</label>

                            <div style={{ display: 'none' }}>
                                <TextField
                                    label="Prefix"
                                    value={formData?.prefix}
                                    onChange={(_, val) => handlePrefixChange(val || "")}
                                />

                            </div>

                            <div style={{ display: "flex", flexDirection: "row", gap: "10px", alignItems: "center" }}>
                                {refFormatData.map((item, index) => (
                                    <div key={index} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "10px" }}>
                                        <span>{item}</span>
                                        <Dropdown
                                            options={[
                                                { key: "Separator", text: "Separator" },
                                                { key: "Concat", text: "Concat" },
                                            ]}
                                            onChange={(e, option) => handleDropdownChange(index, option?.key?.toString() || "Separator")}
                                            selectedKey={customSeparators[index] || "Separator"}
                                        />
                                    </div>
                                ))}
                                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "10px" }}>
                                    <span>{formData?.increment}</span>
                                </div>
                            </div>
                        </div>
                    </div>
                )
            }
        </div>;
    };

    const validation = () => {
        let isValidForm = true;
        const internalName = formData?.TileName?.replace(/[^a-zA-Z0-9]/g, '');
        const isDuplicate = allTiles.filter((item: any) => item.LibraryName === internalName);
        if (formData?.TileName === "" || formData?.TileName === undefined || formData?.TileName === null) {
            setErrors(prevData => ({ ...prevData, TileName: DisplayLabel?.ThisFieldisRequired as string }));
            isValidForm = false;
        }
        else if (isDuplicate.length > 0 && !isEditMode) {
            setErrors(prevData => ({ ...prevData, TileName: DisplayLabel?.TileNameAlreadyExist as string }));
            isValidForm = false;
        }
        else if (formData?.PermissionIds?.length === 0 || formData?.PermissionIds === undefined) {
            setErrors(prevData => ({ ...prevData, Permission: DisplayLabel?.ThisFieldisRequired as string }));
            isValidForm = false;
        }

        else if (formData?.TileAdminId === "" || formData?.TileAdminId === undefined) {
            setErrors(prevData => ({ ...prevData, TileAdmin: DisplayLabel?.ThisFieldisRequired as string }));
            isValidForm = false;
        }


        else if (formData?.isDynamicReference === true) {
            if (refExample === "" || refExample === undefined || refExample === null) {
                setErrors(prevData => ({ ...prevData, refExample: DisplayLabel?.ThisFieldisRequired as string }));
                isValidForm = false;
            }
        }

        else if (formData?.isArchiveAllowed === true) {
            if (formData?.RedundancyData.value === "" || formData?.RedundancyData.value === undefined || formData?.RedundancyData.value === null) {
                setErrors(prevData => ({ ...prevData, RedundancyData: DisplayLabel?.ThisFieldisRequired as string }));
                isValidForm = false;
            }
            else if (formData?.ArchiveVersions === "" || formData?.ArchiveVersions === undefined || formData?.ArchiveVersions === null) {
                setErrors(prevData => ({ ...prevData, ArchiveVersions: DisplayLabel?.ThisFieldisRequired as string }));
                isValidForm = false;
            }
        }

        return isValidForm;
    };

    const submitTileData = () => {
        setErrors({});
        let valid = validation();
        valid ? saveData() : "";
    };


    const UpdateTileData = () => {
        setErrors({});
        let valid = validation();
        valid ? UpdateData() : "";
    };


    const saveData = async () => {
        try {
            setIsDisabled(true);
            // setShowLoader({ display: "block" });
            let ArchiveInternal = "";
            const Internal = formData?.TileName?.replace(/[^a-zA-Z0-9]/g, '');

            if (formData?.isArchiveAllowed == true) {
                ArchiveInternal = formData?.Archive?.replace(/[^a-zA-Z0-9]/g, '');
            }


            const permissionData = formData?.PermissionIds.map((el: any) => ({ Type: "User", IDs: el }));
            permissionData.push({ Type: "Admin", IDs: formData?.TileAdminId }, { Type: "Admin", IDs: admin[0] });


            let option = {
                __metadata: { type: "SP.Data.DMS_x005f_Mas_x005f_TileListItem" },
                TileName: formData?.TileName,
                PermissionId: { results: formData?.PermissionIds },
                TileAdminId: formData?.TileAdminId,
                AllowApprover: formData?.isAllowApprover,
                Active: formData?.isTileStatus,
                IsDynamicReference: formData?.isDynamicReference,
                AllowChildInheritance: formData?.allowChildInheritance,
                Order0: allTiles.length + 1,
                AllowOrder: true,
                ReferenceFormula: refExample,
                Separator: formData?.separator,
                DynamicControl: JSON.stringify(tableData),
                IsArchiveRequired: formData?.isArchiveAllowed,
                ArchiveLibraryName: ArchiveInternal,
                RetentionDays: formData?.RedundancyData?.label === null ? null : parseInt(formData?.RedundancyData?.label),
                ArchiveVersionCount: formData?.ArchiveVersions === null ? null : parseInt(formData?.ArchiveVersions),
                LibraryName: Internal,
                CustomPermission: JSON.stringify(allButtonsWithPermissions),
            };

            const LID = await SaveTileSetting(SiteURL, context.spHttpClient, option);
            if (LID != null) {
                await TileLibrary(context, Internal, LID?.Id, ArchiveInternal, false, tableData, formData?.isArchiveAllowed);
                await breakRoleInheritanceForLib(context, Internal, permissionData);
                setIsOpenEditor(false);
                setIsDisabled(false);
            }
        } catch (error) {
            console.error("Error during save operation:", error);
            setIsDisabled(false);
        }
    };

    const UpdateData = async () => {
        try {
            setIsDisabled(true);
            let ArchiveInternal = "";
            const Internal = formData?.TileName?.replace(/[^a-zA-Z0-9]/g, '');
            createAndUpdateColumn(Internal);

            const permissionData = formData?.PermissionIds.map((el: any) => ({ Type: "User", IDs: el }));
            permissionData.push({ Type: "Admin", IDs: formData?.TileAdminId }, { Type: "Admin", IDs: admin[0] });

            grantPermissionsForLib(context, Internal, permissionData);


            const option = {
                __metadata: { type: "SP.Data.DMS_x005f_Mas_x005f_TileListItem" },
                TileName: formData?.TileName,
                PermissionId: { results: formData?.PermissionIds },
                TileAdminId: formData?.TileAdminId,
                AllowApprover: formData?.isAllowApprover,
                Active: formData?.isTileStatus,
                IsDynamicReference: formData?.isDynamicReference,
                AllowChildInheritance: formData?.allowChildInheritance,
                AllowOrder: true,
                ReferenceFormula: refExample,
                Separator: formData?.separator,
                DynamicControl: JSON.stringify(tableData),
                IsArchiveRequired: formData?.isArchiveAllowed,
                CustomPermission: JSON.stringify(allButtonsWithPermissions),
            };

            await UpdateTileSetting(SiteURL, context.spHttpClient, option, tileID);
            const UpdateData = await getDataById(SiteURL, context.spHttpClient, tileID);
            const UpdateTileID = UpdateData.value;
            if (UpdateTileID != null) {
                if (UpdateTileID[0].IsArchiveRequired === true) {
                    if (UpdateTileID[0].IsArchiveRequired === true) {
                        ArchiveInternal = formData?.Archive.replace(/[^a-zA-Z0-9]/g, '');
                    }
                    else {
                        ArchiveInternal = "";
                    }

                    var items = {
                        __metadata: { type: "SP.Data.DMS_x005f_Mas_x005f_TileListItem" },
                        ArchiveLibraryName: ArchiveInternal,
                        RetentionDays: parseInt(formData?.RedundancyData?.label),
                        ArchiveVersionCount: parseInt(formData?.ArchiveVersions),
                    };
                    await UpdateTileSetting(SiteURL, context.spHttpClient, items, tileID);

                    if (UpdateTileID[0].ArchiveLibraryName === null || UpdateTileID[0].ArchiveLibraryName === undefined || UpdateTileID[0].IsArchiveRequired === false) {
                        await TileLibrary(context, Internal, tileID, ArchiveInternal, true, tableData, formData?.isArchiveAllowed);
                    }
                    else {
                        createAndUpdateColumn(UpdateTileID[0].ArchiveLibraryName);
                    }
                }
            }
            setIsOpenEditor(false);
            setIsDisabled(false);
        }
        catch (error) {
            setIsDisabled(false);
            console.error("Error during save operation:", error);
        }

    };

    const createAndUpdateColumn = async (Internal: string) => {
        const sp = spfi().using(SPFx(context));
        const list = await sp.web.lists.getByTitle(Internal);
        await Promise.all(
            tableData.map(async (item: any) => {
                const isDuplicate = allLibColumn?.some((col: any) => col.InternalName === item.InternalTitleName);

                if (!isDuplicate) {
                    const columnSchema: IColumnSchema = {
                        name: item.InternalTitleName,
                        ColType: getColumnType(item.ColumnType)?.toString(),
                    };
                    await createColumn(list, columnSchema, context);
                }
            })
        );
    };

    const [allLibColumn, setAllLibColumn] = useState([]);
    const getAllColumns = async (TileName: any) => {
        var query = SiteURL + "/_api/web/lists/getbytitle('" + TileName + "')/Fields?$filter=(CanBeDeleted eq true)";
        const response = await GetListData(context, query);
        setAllLibColumn(response.d.results);
    };

    const isAllSelected = (roleTitle: string) => {
        const roleData = allButtonsWithPermissions.find(r => r.Role === roleTitle);

        if (!roleData) return false;

        return Object.values(groupByButtons)
            .flat()
            .every((item: any) =>
                roleData.Permission?.some((p: any) => p.Id === item.Id && p.value)
            );
    };

    const handleSelectAll = (roleTitle: string, isChecked?: boolean) => {
        setAllButtonsWithPermissions((prev: any[]) => {
            return prev.map(role => {
                if (role.Role !== roleTitle) return role;

                const updatedPermissions = Object.values(groupByButtons)
                    .flat()
                    .map((item: any) => ({
                        Id: item.Id,
                        value: isChecked
                    }));

                return {
                    ...role,
                    Permission: updatedPermissions
                };
            });
        });
    };

    const isValidNumberString = (value: string): boolean => {
        return !isNaN(Number(value)) && value.trim() !== "";
    };

    if (isPageLoading) {
        return <PageLoader message="Loading tile form..." minHeight="72vh" />;
    }
    return (
        <>
            <div className="tile-settings-page" data-testid="page-tile-settings">
                <div className="tile-settings-body">
                    <div className="tile-settings-toolbar">
                        <h2 className="tile-settings-subtitle">Workspace Tiles</h2>
                        <div>
                            {!isEditMode ? (
                                <PrimaryButton onClick={submitTileData} text={DisplayLabel?.Submit} className="tile-panel-save-btn" styles={getPrimaryActionButtonStyles(8)} disabled={isDisabled} />
                            ) :
                                <PrimaryButton onClick={UpdateTileData} text={DisplayLabel?.Update} className="tile-panel-save-btn" styles={getPrimaryActionButtonStyles(8)} disabled={isDisabled} />
                            }
                            <DefaultButton
                                className="tile-panel-cancel-btn"
                                styles={getSecondaryActionButtonStyles()}
                                onClick={() => setIsOpenEditor(false)}
                                title="Edit"
                            >{DisplayLabel.Cancel}</DefaultButton>
                        </div>
                    </div>
                    <div className="tile-settings-table-wrap" data-testid="table-tiles">
                        <div className="tile-panel-body" data-testid="container-edit-tile-panel">
                            <div className="tile-panel-section">
                                <div
                                    className="tile-panel-section-header"
                                    onClick={() => toggleSection('tileDetails')}
                                    data-testid="toggle-section-tileDetails"
                                >
                                    <span className="tile-panel-section-title">{DisplayLabel.TileDetails}</span>
                                    {expandedSections.has('tileDetails') ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
                                </div>
                                {expandedSections.has('tileDetails') && (
                                    <div className="tile-panel-section-content">
                                        <div className="grid-2">
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.TileName}<span className="tile-form-required">*</span></label>
                                                <TextField
                                                    value={formData?.TileName}
                                                    onChange={(_, val) => setFormData((prev: any) => ({ ...prev, TileName: val || '', Archive: `Archive ${val}` }))}
                                                    className="tile-form-input"
                                                    data-testid="input-tile-name"
                                                    errorMessage={errors?.TileName}
                                                />
                                            </div>

                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.TileAdmin1}<span className="tile-form-required">*</span></label>
                                                <PeoplePicker
                                                    context={peoplePickerContext}
                                                    personSelectionLimit={1}
                                                    showtooltip={true}
                                                    showHiddenInUI={false}
                                                    ensureUser={true}
                                                    principalTypes={[PrincipalType.User]}
                                                    onChange={(users: any[]) => {
                                                        setFormData((prevValues) => ({
                                                            ...prevValues,
                                                            TileAdminId: users[0].id,
                                                            TileAdminEmail: users[0].email
                                                        }));

                                                    }}
                                                    defaultSelectedUsers={formData?.TileAdminEmail}
                                                />
                                                <FieldError message={errors.TileAdmin} />
                                            </div>
                                        </div>
                                        <div className="tile-form-field">
                                            <label className="tile-form-label">{DisplayLabel.AccessToTile}<span className="tile-form-required">*</span></label>
                                            <PeoplePicker
                                                context={peoplePickerContext}
                                                personSelectionLimit={20}
                                                showtooltip={true}
                                                showHiddenInUI={false}
                                                ensureUser={true}
                                                principalTypes={[PrincipalType.User, PrincipalType.DistributionList, PrincipalType.SecurityGroup, PrincipalType.SharePointGroup]}
                                                onChange={(users: any[]) => {
                                                    const ids = users.map((user: any) => user.id || user.Id);
                                                    const emails = users.map((user: any) => user.secondaryText || user.loginName || user.email);

                                                    setFormData((prevValues) => ({
                                                        ...prevValues,
                                                        PermissionIds: ids,
                                                        PermissionEmail: emails
                                                    }));
                                                }}
                                                defaultSelectedUsers={formData?.PermissionEmail}
                                            />
                                            <FieldError message={errors.Permission} />
                                        </div>
                                        <div className="grid-3">
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.TileStatus}</label>
                                                <Toggle
                                                    checked={formData?.isTileStatus}
                                                    onChange={(_, checked) => setFormData((prev: any) => ({ ...prev, isTileStatus: !!checked }))}
                                                    className="tile-form-toggle"
                                                    data-testid="toggle-tile-status"
                                                />
                                            </div>
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.AllowApprover}</label>
                                                <Toggle
                                                    checked={formData?.isAllowApprover}
                                                    onChange={(_, checked) => setFormData((prev: any) => ({ ...prev, isAllowApprover: !!checked }))}
                                                    className="tile-form-toggle"
                                                    data-testid="toggle-allow-approver"
                                                />
                                            </div>
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.AllowChildInheritance}</label>
                                                <Toggle
                                                    checked={formData?.allowChildInheritance}
                                                    onChange={(_, checked) => setFormData((prev: any) => ({ ...prev, allowChildInheritance: !!checked }))}
                                                    className="tile-form-toggle"
                                                    data-testid="toggle-child-inheritance"
                                                />
                                            </div>
                                        </div>
                                    </div>
                                )}
                            </div>

                            <div className="tile-panel-section">
                                <div
                                    className="tile-panel-section-header"
                                    onClick={() => toggleSection('fields')}
                                    data-testid="toggle-section-fields"
                                >
                                    <span className="tile-panel-section-title">{DisplayLabel.Field}</span>
                                    {expandedSections.has('fields') ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
                                </div>
                                {expandedSections.has('fields') && (
                                    <div className="tile-panel-section-content">
                                        {renderFieldTable()}
                                    </div>
                                )}
                            </div>

                            <div className="tile-panel-section">
                                <div
                                    className="tile-panel-section-header"
                                    onClick={() => toggleSection('buttons')}
                                    data-testid="toggle-section-fields"
                                >
                                    <span className="tile-panel-section-title">Buttons Permissions</span>
                                    {expandedSections.has('buttons') ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
                                </div>
                                {expandedSections.has('buttons') && (
                                    <div className="tile-panel-section-content">
                                        <table className="fluent-table">
                                            <thead>
                                                <tr>
                                                    <th></th>
                                                    <th>{DisplayLabel?.Role}</th>
                                                    {roles && roles.map((role: any) => (<th>{role.Title}</th>))}
                                                </tr>
                                                <tr>
                                                    <td></td>
                                                    <td></td>
                                                    {roles && roles.map((role: any) => (
                                                        <td>
                                                            <PeoplePicker context={peoplePickerContext}
                                                                personSelectionLimit={20}
                                                                showtooltip={true}
                                                                required={true}
                                                                errorMessage={errors?.AccessTileUserErr}
                                                                ensureUser={true}
                                                                onChange={async (items: any[]) => {
                                                                    const userIds = await Promise.all(
                                                                        items.map(async (item: any) => {
                                                                            let userid: number = 0;
                                                                            if (isValidNumberString(item.id)) {
                                                                                userid = item;
                                                                            } else {
                                                                                const data = await getUserIdFromLoginName(context, item.id);
                                                                                userid = data;
                                                                            };
                                                                            return userid;
                                                                        })
                                                                    );
                                                                    grantButtonPermission(role.Title, userIds);
                                                                }}
                                                                showHiddenInUI={false}
                                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup]}
                                                                defaultSelectedUsers={isEditMode ? bindPermission(role.Title) : undefined}
                                                                styles={{ root: { order: -1 } }}
                                                            />
                                                        </td>
                                                    ))}
                                                </tr>
                                            </thead>

                                            <tbody>
                                                <tr>
                                                    <td>Select All</td>
                                                    <td></td>
                                                    {roles && roles.map((role: any) => (
                                                        <td>
                                                            <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
                                                                <Checkbox
                                                                    checked={isAllSelected(role.Title)}
                                                                    onChange={(_, val) => handleSelectAll(role.Title, !!val)}
                                                                />
                                                            </div>
                                                        </td>
                                                    ))}
                                                </tr>
                                                {Object.keys(groupByButtons).map((group: any) => (
                                                    <React.Fragment key={group}>
                                                        {groupByButtons[group].map((item: any, index: number) => (
                                                            <tr key={item.Id} style={{ borderBottom: "1px solid #eee" }}>
                                                                {index === 0 && <td rowSpan={groupByButtons[group].length}> {group} </td>}
                                                                <td style={{ padding: "8px 12px" }}>{item.Title}</td>
                                                                {roles.map((role: any) => {
                                                                    const foundRole = allButtonsWithPermissions.find(
                                                                        (r) => r.Role === role.Title
                                                                    );
                                                                    const foundPerm = foundRole?.Permission?.find(
                                                                        (p: any) => p.Id === item.Id
                                                                    );

                                                                    return (
                                                                        <td key={`${role.Id}_${item.Id}`}>
                                                                            <div style={{ display: "flex", justifyContent: "center", width: "100%" }}>
                                                                                <Checkbox
                                                                                    checked={!!foundPerm?.value}
                                                                                    onChange={(_, val) =>
                                                                                        handleCheckboxChangeButton(role.Title, item.Id, !!val)
                                                                                    }
                                                                                />
                                                                            </div>
                                                                        </td>
                                                                    );
                                                                })}
                                                            </tr>
                                                        ))}
                                                    </React.Fragment>
                                                ))}
                                            </tbody>

                                        </table>
                                    </div>
                                )}
                            </div>

                            <div className="tile-panel-section">
                                <div
                                    className="tile-panel-section-header"
                                    onClick={() => toggleSection('referenceNo')}
                                    data-testid="toggle-section-referenceNo"
                                >
                                    <span className="tile-panel-section-title">{DisplayLabel.ReferenceNoDetails}</span>
                                    {expandedSections.has('referenceNo') ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
                                </div>
                                {expandedSections.has('referenceNo') && (
                                    <div className="tile-panel-section-content">
                                        <div className="tile-form-grid tile-form-grid-2">
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.IsDynamicReference}<span className="tile-form-required">*</span></label>
                                                <Toggle
                                                    checked={formData?.isDynamicReference}
                                                    onChange={(_, checked) => {
                                                        setFormData((prev: any) => ({ ...prev, isDynamicReference: !!checked }));
                                                        !checked && setRefExample(refrenceNOData);
                                                    }}
                                                    className="tile-form-toggle"
                                                    data-testid="toggle-dynamic-reference"
                                                />
                                            </div>
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">{DisplayLabel.DefaultReferenceExample}</label>
                                                <TextField
                                                    value={refExample}
                                                    className="tile-form-input"
                                                    data-testid="input-reference-example"
                                                    readOnly
                                                />
                                            </div>
                                        </div>
                                        {renderRefForm()}
                                    </div>
                                )}
                            </div>

                            <div className="tile-panel-section">
                                <div
                                    className="tile-panel-section-header"
                                    onClick={() => toggleSection('archive')}
                                    data-testid="toggle-section-archive"
                                >
                                    <span className="tile-panel-section-title">Archive Section</span>
                                    {expandedSections.has('archive') ? <ChevronUp20Regular /> : <ChevronDown20Regular />}
                                </div>
                                {expandedSections.has('archive') && (
                                    <div className="tile-panel-section-content">
                                        <div className="grid-2">
                                            <div className="tile-form-field">
                                                <label className="tile-form-label">Is Archive Allowed<span className="tile-form-required">*</span></label>
                                                <Toggle
                                                    checked={formData?.isArchiveAllowed}
                                                    onChange={(_, checked) => setFormData((prev: any) => ({ ...prev, isArchiveAllowed: !!checked }))}
                                                    className="tile-form-toggle"
                                                    data-testid="toggle-archive-allowed"
                                                />
                                            </div>
                                            {formData?.isArchiveAllowed && (
                                                <>
                                                    <div className="column6">
                                                        <label>{DisplayLabel?.ArchiveDocumentLibraryName}</label>
                                                        <TextField
                                                            placeholder=" "
                                                            value={formData.Archive}
                                                            disabled
                                                        />
                                                    </div>
                                                    <div className="column6">
                                                        <label style={{ display: 'block' }}>{DisplayLabel?.SelectArchiveDays}<span style={{ color: "red" }}>*</span></label>
                                                        <Select
                                                            options={redundancyData}
                                                            value={formData?.RedundancyData || {}}
                                                            onChange={handleArchiveDropdownChange}
                                                            isSearchable
                                                            placeholder={DisplayLabel?.Selectanoption}
                                                            menuPortalTarget={document.body}
                                                            menuPosition="fixed"
                                                            styles={{
                                                                menuPortal: (base: CSSObjectWithLabel) => ({ ...base, zIndex: 9999 }),
                                                                menu: (base: CSSObjectWithLabel) => ({ ...base, zIndex: 9999 })
                                                            }}
                                                            errorMessage={errors?.RedundancyData}
                                                        />
                                                        <FieldError message={errors?.RedundancyData} />
                                                    </div>
                                                    <div className="column6">
                                                        <label style={{ display: 'block' }}>{DisplayLabel?.ArchiveVersions}<span style={{ color: "red" }}>*</span></label>
                                                        <TextField
                                                            placeholder=" "
                                                            value={formData?.ArchiveVersions}
                                                            errorMessage={errors?.ArchiveVersions}
                                                            onChange={(el, value) => setFormData(prevData => ({ ...prevData, ArchiveVersions: value }))}
                                                        />
                                                    </div>
                                                </>
                                            )}
                                        </div>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </>
    );
};

export default React.memo(TileForm);
