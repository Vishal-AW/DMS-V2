import * as React from 'react';
import { CheckmarkCircle24Filled, Warning24Filled } from "@fluentui/react-icons";
import {
    Dialog,
    DialogType,
    DialogFooter,
    PrimaryButton,
    DefaultButton,
    Layer,
} from "@fluentui/react";
import './PopupStyle.css';


interface IPopupboxProps {
    isPopupBoxVisible: boolean;
    hidePopup: () => void;
    msg: string;
    type?: "success" | "warning" | "insert" | "checkin" | "checkout" | "approve" | "reject" | "delete" | "update" | "restore" | "grant" | "remove";
}

const getPopupContent = (msg: string, type: NonNullable<IPopupboxProps["type"]>) => {
    const normalizedMsg = (msg || "").toLowerCase();

    if (type === "warning") {
        return {
            title: "Attention",
            message: msg
        };
    }

    if (type === "insert") {
        return {
            title: "Request Submitted",
            message: "The request has been submitted successfully."
        };
    }

    if (type === "checkin") {
        return {
            title: "Check-In Complete",
            message: "The document has been checked in successfully."
        };
    }

    if (type === "checkout") {
        return {
            title: "Check-Out Complete",
            message: "The document has been checked out successfully."
        };
    }

    if (type === "approve") {
        return {
            title: "Request Approved",
            message: "The request has been approved successfully."
        };
    }

    if (type === "reject") {
        return {
            title: "Request Rejected",
            message: "The request has been rejected successfully."
        };
    }

    if (type === "delete") {
        return {
            title: "Deleted Successfully",
            message: "The item has been deleted successfully."
        };
    }

    if (type === "update") {
        return {
            title: "Updated Successfully",
            message: "The changes have been saved successfully."
        };
    }

    if (type === "restore") {
        return {
            title: "Restored Successfully",
            message: "The item has been restored successfully."
        };
    }

    if (type === "grant") {
        return {
            title: "Access Granted",
            message: "Permissions have been granted successfully."
        };
    }

    if (type === "remove") {
        return {
            title: "Access Removed",
            message: "Permissions have been removed successfully."
        };
    }

    if (normalizedMsg.includes("submit")) {
        return {
            title: "Request Submitted",
            message: "The request has been submitted successfully."
        };
    }

    if (normalizedMsg.includes("checkin") || normalizedMsg.includes("check in")) {
        return {
            title: "Check-In Complete",
            message: "The document has been checked in successfully."
        };
    }

    if (normalizedMsg.includes("checkout") || normalizedMsg.includes("check out")) {
        return {
            title: "Check-Out Complete",
            message: "The document has been checked out successfully."
        };
    }

    if (normalizedMsg.includes("approve")) {
        return {
            title: "Request Approved",
            message: "The request has been approved successfully."
        };
    }

    if (normalizedMsg.includes("reject")) {
        return {
            title: "Request Rejected",
            message: "The request has been rejected successfully."
        };
    }

    if (normalizedMsg.includes("delete")) {
        return {
            title: "Deleted Successfully",
            message: "The item has been deleted successfully."
        };
    }

    if (normalizedMsg.includes("update")) {
        return {
            title: "Updated Successfully",
            message: "The changes have been saved successfully."
        };
    }

    if (normalizedMsg.includes("restore")) {
        return {
            title: "Restored Successfully",
            message: "The item has been restored successfully."
        };
    }

    if (normalizedMsg.includes("grant")) {
        return {
            title: "Access Granted",
            message: "Permissions have been granted successfully."
        };
    }

    if (normalizedMsg.includes("remove")) {
        return {
            title: "Access Removed",
            message: "Permissions have been removed successfully."
        };
    }

    return {
        title: "Success",
        message: msg
    };
};

const PopupBox: React.FC<IPopupboxProps> = ({ isPopupBoxVisible, hidePopup, msg, type = "success", }) => {
    if (!isPopupBoxVisible) {
        return null;
    }

    const isSuccess = type !== "warning";
    const accentClass = isSuccess ? "alert-success" : "alert-warning";
    const { title, message } = getPopupContent(msg, type);

    return (

        <>{isPopupBoxVisible ?
            <Layer>
                <div className="backdrop"></div>
                <div className={`alert ${accentClass}`}>
                    <div className="alert-content">
                        <div className="alert-icon">
                            {isSuccess ? (
                                <CheckmarkCircle24Filled
                                    style={{ color: "#22c55e", fontSize: 76 }}
                                />
                            ) : (
                                <Warning24Filled
                                    style={{ color: "#f59e0b", fontSize: 76 }}
                                />
                            )}
                        </div>
                        <h3 className="alert-title">{title}</h3>
                        <p className="alert-message">{message}</p>
                        <button className="alert-button" onClick={hidePopup}>OK</button>
                    </div>
                </div>
            </Layer>
            : <></>
        }
        </>


    );
};

export default React.memo(PopupBox);




interface IConfirmboxProps {
    hideDialog: boolean;
    closeDialog: () => void;
    handleConfirm?: (value: boolean) => void;
    handleNo?: () => void;
    msg: string;
    Yes?: string;
    No?: string;
}
export const ConfirmationDialog: React.FC<IConfirmboxProps> = ({ hideDialog, closeDialog, handleConfirm, handleNo, msg, Yes = "Yes", No = "No" }) => {


    const onYesClick = () => {
        if (handleConfirm) {
            handleConfirm(true);
        }
        closeDialog();
    };

    const onNoClick = () => {
        if (handleConfirm) {
            handleConfirm(false);
        }

        if (handleNo) {
            handleNo();
        } else {
            closeDialog();
        }
    };

    return (
        <div>
            <Dialog
                hidden={!hideDialog}
                onDismiss={closeDialog}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: "Confirm Action",
                    subText: msg,
                }}
                modalProps={{
                    isBlocking: true,
                    isDarkOverlay: false
                }}
            >
                <DialogFooter>

                    {/* <PrimaryButton text="Yes" onClick={() => handleConfirm(true)} /> */}
                    {/* <PrimaryButton text={Yes} style={{ backgroundColor: '#ca5010', border: '1px solid #ca5010' }} onClick={() => handleConfirm(true)} />
                    <DefaultButton text={No} onClick={closeDialog} /> */}
                    <PrimaryButton
                        text={Yes}
                        style={{ backgroundColor: '#ca5010', border: '1px solid #ca5010' }}
                        onClick={onYesClick}
                    />

                    <DefaultButton
                        text={No}
                        onClick={onNoClick}
                    />
                </DialogFooter>
            </Dialog>
        </div>
    );
};
