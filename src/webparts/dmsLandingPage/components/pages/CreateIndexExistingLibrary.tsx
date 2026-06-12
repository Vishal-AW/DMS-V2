import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";

interface IProps {
    context: WebPartContext;
}

const IndexLibraryButton: React.FC<IProps> = ({ context }) => {

    const [isLoading, setIsLoading] = React.useState(false);
    const [status, setStatus] = React.useState("");

    const handleIndexClick = async () => {
        const libraryTitle = "Your Library Name"; // 🔁 replace with your library name
        setIsLoading(true);
        setStatus("");

        const sp = spfi().using(SPFx(context));
        const columnsToIndex = ["FSObjType", "FileLeafRef", "FileDirRef", "Modified", "Created"];
        let successCount = 0;

        try {
            const list = sp.web.lists.getByTitle(libraryTitle);

            for (const colName of columnsToIndex) {
                try {
                    await list.fields
                        .getByInternalNameOrTitle(colName)
                        .update({ Indexed: true });
                    successCount++;
                } catch {
                    console.warn(`Skipped: ${colName}`);
                }
            }

            setStatus(`✅ Done! ${successCount}/${columnsToIndex.length} columns indexed.`);

        } catch (err) {
            setStatus("❌ Error: Library not found or access denied.");
        } finally {
            setIsLoading(false);
        }
    };

    return (
        <div>
            <button
                onClick={handleIndexClick}
                disabled={isLoading}
                style={{
                    padding: "8px 16px",
                    backgroundColor: isLoading ? "#ccc" : "#0078d4",
                    color: "white",
                    border: "none",
                    borderRadius: "4px",
                    cursor: isLoading ? "not-allowed" : "pointer"
                }}
            >
                {isLoading ? "Indexing..." : "Index Library Columns"}
            </button>

            {status && (
                <p style={{ marginTop: "8px", color: status.includes("❌") ? "red" : "green" }}>
                    {status}
                </p>
            )}
        </div>
    );
};

export default IndexLibraryButton;