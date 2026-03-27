import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

interface IPageLoaderProps {
    message?: string;
    minHeight?: string;
}

const PageLoader: React.FC<IPageLoaderProps> = ({
    message = "Loading...",
    minHeight = "60vh"
}) => {
    return (
        <div
            style={{
                minHeight,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                padding: "24px"
            }}
        >
            <div
                style={{
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    gap: "12px",
                    padding: "28px 32px",
                    borderRadius: "18px",
                    background: "linear-gradient(180deg, #ffffff 0%, #f7fafc 100%)",
                    border: "1px solid #dbe7f3",
                    boxShadow: "0 14px 40px rgba(15, 108, 189, 0.08)"
                }}
            >
                <Spinner size={SpinnerSize.large} label={message} />
            </div>
        </div>
    );
};

export default React.memo(PageLoader);
