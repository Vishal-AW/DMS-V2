import * as React from 'react';
import { CheckmarkCircle24Filled } from "@fluentui/react-icons";

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
    type?: "success" | "warning";
}

const PopupBox: React.FC<IPopupboxProps> = ({ isPopupBoxVisible, hidePopup, msg, type = "success", }) => {
    console.log("Alert Box");
    if (!isPopupBoxVisible) {
        return null;
    }

    const isSuccess = type === "success";

    return (

        <>{isPopupBoxVisible ?
            <Layer>
                <div className="backdrop"></div>
                <div className="alert">
                    <div className="alert-content">
                        <div className="alert-icon">
                            <CheckmarkCircle24Filled
                                style={{ color: "#50cd89", fontSize: 80 }}
                            />
                        </div>
                        <p className="alert-message">{msg}</p>
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
