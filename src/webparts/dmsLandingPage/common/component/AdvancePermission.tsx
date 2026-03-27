import { Checkbox, DefaultButton, Panel, PanelType, PrimaryButton } from "@fluentui/react";
import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPeoplePickerContext, PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { useCallback, useEffect, useState } from "react";
import { commonPostMethod, getPermission } from "../../../../Services/GeneralDocument";
import PopupBox, { ConfirmationDialog } from "./PopupBox";
import { ILabel } from "../../../../Intrface/ILabel";
import Select from "react-select";
import { Field } from "@fluentui/react-components";
import PageLoader from "./PageLoader";
import FieldError from "./FieldError";
export interface IAdvanceProps {
    isOpen: boolean;
    dismissPanel: () => void;
    context: WebPartContext;
    LibraryName: string;
    folderId: number;
}

const AdvancePermission: React.FC<IAdvanceProps> = ({ isOpen, dismissPanel, context, folderId, LibraryName }) => {

    const [option, setOption] = useState<string | null>(null);
    const [hasUniquePermission, setHasUniquePermission] = useState<boolean>(false);
    const [userData, setUserData] = useState<any[]>([]);
    const [hideDialog, setHideDialog] = useState<boolean>(false);
    const [isCheckedUser, setIsCheckedUser] = useState<string[]>([]);
    const [isPopupBoxVisible, setIsPopupBoxVisible] = useState<boolean>(false);
    const [popupType, setPopupType] = useState<"success" | "warning" | "insert" | "checkin" | "checkout" | "approve" | "reject" | "delete" | "update" | "restore" | "grant" | "remove">("success");
    const [message, setMessage] = useState<string>("");
    const [selectedUser, setSelectedUser] = useState<any[]>([]);
    const [selectedUserError, setSelectedUserError] = useState("");

    const [peoplePickerKey, setPeoplePickerKey] = useState(0);
    const [selectedPermissionError, setSelectedPermissionError] = useState("");
    const DisplayLabel: ILabel = JSON.parse(localStorage.getItem('DisplayLabel') || '{}');
    const [alertMsg, setAlertMsg] = useState("");
    const [isLoading, setIsLoading] = useState(false);
    const peoplePickerRef = React.useRef<any>(null);
    const peoplePickerContext: IPeoplePickerContext = {
        absoluteUrl: context.pageContext.web.absoluteUrl,
        msGraphClientFactory: context.msGraphClientFactory as any,
        spHttpClient: context.spHttpClient as any
    };




    const permissionDetails: Record<string, string> = {
        "1073741829": DisplayLabel.FullControlAccessDec,
        "1073741830": DisplayLabel.EditAccessDec,
        "1073741826": DisplayLabel.ReadAccessDec,
    };

    const clearGrantFields = () => {
        setSelectedUser([]);
        setOption(null);
        setSelectedUserError("");
        setSelectedPermissionError("");
        setPeoplePickerKey(prev => prev + 1);
    };

    useEffect(() => {
        if (isOpen) bindPermission();

    }, [isOpen]);

    const bindPermission = async () => {
        if (isOpen) {
            setIsLoading(true);
            setIsCheckedUser([]);
            try {

                const checkUniquePermissionQuery = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LibraryName}')/items(${folderId})/HasUniqueRoleAssignments`;
                const getMemberQuery = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LibraryName}')/items(${folderId})/roleassignments?$expand=RoleDefinitionBindings,Member`;

                const uniquePermissionResponse = await getPermission(checkUniquePermissionQuery, context);
                setHasUniquePermission(uniquePermissionResponse.value);

                const memberDataResponse = await getPermission(getMemberQuery, context);
                setUserData(memberDataResponse?.value || []);
                clearGrantFields();

            } catch (error) {
                console.error("Error binding permissions: ", error);
            } finally {
                setIsLoading(false);
            }
        }
        //  setSelectedUser([]);
        //  setOption("null");
    };

    const handleSelectAllChange = () => {
        if (isCheckedUser.length === userData.length) {
            setIsCheckedUser([]); // Uncheck all
        } else {
            setIsCheckedUser(userData.map((user: any) => user.Member.Id)); // Check all
        }
    };

    const handleCheckboxChange = (userId: string) => {
        setIsCheckedUser((prev) =>
            prev.includes(userId)
                ? prev.filter((id) => id !== userId) // Remove userId if already checked
                : [...prev, userId] // Add userId if not checked
        );
    };

    const closeDialog = useCallback(() => setHideDialog(false), []);

    const handleConfirm = useCallback(

        async (value: boolean) => {
            if (value) {

                setHideDialog(false);
                const requestUri = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LibraryName}')/items(${folderId})/breakroleinheritance(true)`;
                try {
                    await commonPostMethod(requestUri, context);
                    setAlertMsg(DisplayLabel.StopInheritingSuccessMsg);
                    setPopupType("update");
                    setIsPopupBoxVisible(true);
                    bindPermission();
                } catch (error) {
                    console.error("Error stopping inheritance: ", error);
                }
            }
        },
        [context, folderId, LibraryName]
    );

    const hidePopup = useCallback(() => setIsPopupBoxVisible(false), []);

    const removeUserPermission = async () => {
        let count = 0;
        for (const userId of isCheckedUser) {
            const requestUri = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LibraryName}')/items(${folderId})/roleassignments/removeroleassignment(principalid=${userId})`;
            try {
                setAlertMsg(DisplayLabel.AccessHasRemoved);
                setPopupType("remove");
                await commonPostMethod(requestUri, context);
                count++;
                if (count === isCheckedUser.length) setIsPopupBoxVisible(true);
                bindPermission();
            } catch (error) {
                console.error("Error removing user permission: ", error);
            }
        }
    };
    const grantPermission = async () => {
        let isValid = true;

        setSelectedPermissionError("");
        setSelectedUserError("");

        if (!selectedUser || selectedUser.length === 0) {
            setSelectedUserError(DisplayLabel.ThisFieldisRequired);
            isValid = false;
        }

        if (!option) {
            setSelectedPermissionError(DisplayLabel.ThisFieldisRequired);
            isValid = false;
        }

        if (!isValid) {
            return;
        }

        try {
            await Promise.all(
                selectedUser.map((userId: any) => {
                    const requestUri = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${LibraryName}')/items(${folderId})/roleassignments/addroleassignment(principalid=${userId},roledefid=${option})`;
                    return commonPostMethod(requestUri, context);
                })
            );

            setAlertMsg(DisplayLabel.AccessHasGranted);
            setPopupType("grant");
            setIsPopupBoxVisible(true);
            clearGrantFields();
            await bindPermission();
        } catch (error) {
            console.error("Error granting permissions: ", error);
        }
    };

    const otions = [
        { value: "1073741829", label: DisplayLabel.FullControlAccess },
        { value: "1073741830", label: DisplayLabel.EditAccess },
        { value: "1073741826", label: DisplayLabel.ReadAccess },
    ];




    return (
        <div>
            <Panel
                headerText={DisplayLabel.AdvancePermission}
                isOpen={isOpen}
                onDismiss={() => {
                    if (!hideDialog && !isPopupBoxVisible) {
                        dismissPanel();
                        clearGrantFields();

                    }
                }}
                isBlocking={true}
                closeButtonAriaLabel="Close"
                type={PanelType.medium}
            >
                {isLoading ? <PageLoader message="Loading permissions..." minHeight="40vh" /> :
                    <div>

                        <div className="grid-2">
                            <div className="col-md-6">
                                <DefaultButton
                                    style={{
                                        backgroundColor: hasUniquePermission ? '#f3f2f1' : '#ca5010',
                                        border: `1px solid ${hasUniquePermission ? '#f3f2f1' : '#ca5010'}`,
                                        color: hasUniquePermission ? '#a19f9d' : '#ffffff',
                                        cursor: hasUniquePermission ? 'not-allowed' : 'pointer',
                                    }}
                                    text={DisplayLabel.StopInheritingPermission}
                                    disabled={hasUniquePermission}
                                    onClick={() => {
                                        setMessage(DisplayLabel.StopInheritingConfirmMsg);
                                        setHideDialog(true);
                                        dismissPanel();
                                    }}
                                />
                            </div>
                            <div className="col-md-6">
                                <PrimaryButton
                                    style={{
                                        backgroundColor: isCheckedUser.length > 0 ? '#ca5010' : undefined,
                                        border: isCheckedUser.length > 0 ? '1px solid #ca5010' : undefined,
                                    }}
                                    text={DisplayLabel.RemoveUserPermission}
                                    disabled={isCheckedUser.length === 0}
                                    onClick={removeUserPermission}
                                />
                            </div>
                        </div>

                        <div className="grid-2">
                            <Field label={DisplayLabel?.EnterName} required>

                                <PeoplePicker
                                    key={peoplePickerKey}
                                    context={peoplePickerContext}
                                    personSelectionLimit={20}
                                    showtooltip={true}
                                    required={true}
                                    ensureUser={true}
                                    onChange={
                                        async (items) => {
                                            try {

                                                const userIds = items.map(user => user.id) || [];
                                                setSelectedUser(userIds);
                                                setSelectedUserError("");
                                            } catch (error) {
                                                console.error("Error fetching user IDs:", error);
                                            }
                                        }}

                                    ref={peoplePickerRef.current?.clear()}
                                    principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup, PrincipalType.SecurityGroup]}
                                />
                                <FieldError message={selectedUserError} />
                            </Field>

                            <Field label={DisplayLabel?.SelectPermissionLevel} required>
                                <Select
                                    required
                                    options={otions}
                                    value={otions.find((item: any) => item.value === option) || null}
                                    onChange={(opt: any) => {
                                        setOption(opt?.value as string ?? null);
                                        setSelectedPermissionError("");
                                    }}
                                    isSearchable
                                    isClearable
                                    placeholder={DisplayLabel?.Selectanoption}
                                    style={{ margintop: "7px" }}
                                />
                                <FieldError message={selectedPermissionError} />
                            </Field>
                        </div>

                        <div className="row">
                            <div className="col-md-12">
                                {option && <span style={{ color: "red" }}>Note: {permissionDetails[option]}</span>}
                            </div>
                        </div>
                        <div className="row">
                            <PrimaryButton text={DisplayLabel.GrantPermissions} onClick={grantPermission} className="workspace-new-request-btn" />
                        </div>

                        <div className="row">
                            <table className="fluent-table">
                                <thead>
                                    <tr>
                                        <th style={{ width: 50 }}>
                                            <Checkbox
                                                checked={isCheckedUser.length === userData.length}
                                                onChange={handleSelectAllChange}
                                            />
                                        </th>
                                        <th>Name</th>
                                        <th>Permission Levels</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {userData.map((el: any) => (
                                        <tr key={el.Id}>
                                            <td>
                                                {el.Member.Title !== "ProjectAdmin" && (
                                                    <Checkbox
                                                        checked={isCheckedUser.includes(el.Member.Id)}
                                                        onChange={() => handleCheckboxChange(el.Member.Id)}
                                                    />
                                                )}
                                            </td>
                                            <td>{el.Member.Title}</td>
                                            <td>
                                                {el.RoleDefinitionBindings.map((item: any) => (
                                                    <React.Fragment key={item.Id}>
                                                        <p>{item.Name}</p>
                                                        <p >{item.Description}</p>
                                                    </React.Fragment>
                                                ))}
                                            </td>
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    </div>
                }
            </Panel>
            <ConfirmationDialog hideDialog={hideDialog} closeDialog={closeDialog} handleConfirm={handleConfirm} msg={message} />
            <PopupBox isPopupBoxVisible={isPopupBoxVisible} hidePopup={hidePopup} msg={alertMsg} type={popupType} />
        </div>
    );
};

export default React.memo(AdvancePermission);
