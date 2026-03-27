import * as React from "react";
import { Info12Regular } from "@fluentui/react-icons";

interface IFieldErrorProps {
    message?: string;
}

const FieldError: React.FC<IFieldErrorProps> = ({ message }) => {
    if (!message) {
        return null;
    }

    return (
        <div className="form-error-message" role="alert">
            <Info12Regular className="form-error-icon" />
            <span>{message}</span>
        </div>
    );
};

export default React.memo(FieldError);
