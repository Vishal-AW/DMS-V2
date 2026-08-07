/* eslint-disable */
import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import {
    TextField,
    Panel,
    PanelType,
    DefaultButton,
    PrimaryButton,
    Toggle
} from "@fluentui/react";
import { Field } from "@fluentui/react-components";
import Select from "react-select";
import { getPrimaryActionButtonStyles, getSecondaryActionButtonStyles } from "../../common/component/buttonStyles";
import { SaveFolderMaster, UpdateFolderMaster, getFoldersByTemplateId } from "../../../../Services/FolderMasterService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface AddFolderPanelProps {
    context: WebPartContext;
    isOpen: boolean;
    onDismiss: () => void;
    onSaved: () => void;
    templateId: number;
    templateName: string;
    editFolderData?: any | null;
}

const AddFolderPanel: React.FC<AddFolderPanelProps> = ({
    context,
    isOpen,
    onDismiss,
    onSaved,
    templateId,
    templateName,
    editFolderData = null
}) => {
    const [folderName, setFolderName] = useState("");
    const [isChildFolder, setIsChildFolder] = useState(false);
    const [active, setActive] = useState(true);
    const [parentFolderId, setParentFolderId] = useState<number | undefined>();
    const [parentFolderOptions, setParentFolderOptions] = useState<any[]>([]);

    const [nameError, setNameError] = useState("");
    const [parentError, setParentError] = useState("");

    const isEditMode = editFolderData !== null;

    useEffect(() => {
        if (isOpen) {
            if (editFolderData) {
                // Edit mode - populate fields
                setFolderName(editFolderData.FolderName || "");
                setActive(editFolderData.Active !== false);
                setIsChildFolder(!!editFolderData.ParentFolderIdId);
                setParentFolderId(editFolderData.ParentFolderIdId || undefined);
            } else {
                // Add mode - reset fields
                setFolderName("");
                setActive(true);
                setIsChildFolder(false);
                setParentFolderId(undefined);
            }
            setNameError("");
            setParentError("");

            // Load parent folder options for this template
            loadParentFolders();
        }
    }, [isOpen, editFolderData, templateId]);

    const loadParentFolders = useCallback(async () => {
        try {
            const res: any = await getFoldersByTemplateId(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                templateId
            );
            const folders = res?.value || [];
            // Filter out the current folder being edited (to prevent self-reference)
            const filtered = editFolderData
                ? folders.filter((f: any) => f.ID !== editFolderData.ID)
                : folders;
            setParentFolderOptions(
                filtered.map((item: any) => ({
                    key: item.ID,
                    text: item.FolderName
                }))
            );
        } catch (error) {
            console.error("Error loading parent folders:", error);
            setParentFolderOptions([]);
        }
    }, [context, templateId, editFolderData]);

    const handleSave = async () => {
        // Validation
        if (!folderName.trim()) {
            setNameError("Folder Name is required");
            return;
        }
        if (!/^[a-zA-Z0-9 ]+$/.test(folderName.trim())) {
            setNameError("Special characters are not allowed");
            return;
        }
        if (isChildFolder && !parentFolderId) {
            setParentError("Parent Folder is required");
            return;
        }

        setNameError("");
        setParentError("");

        const option: any = {
            FolderName: folderName.trim(),
            Active: active,
            TemplateNameId: templateId,
            IsParentFolder: isChildFolder,
            ParentFolderIdId: isChildFolder ? (parentFolderId ?? null) : null
        };

        try {
            if (isEditMode && editFolderData?.ID) {
                await UpdateFolderMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option,
                    editFolderData.ID
                );
            } else {
                await SaveFolderMaster(
                    context.pageContext.web.absoluteUrl,
                    context.spHttpClient,
                    option
                );
            }
            onSaved();
            onDismiss();
        } catch (error) {
            console.error("Error saving folder:", error);
        }
    };

    const parentFolderSelectOptions = parentFolderOptions.map((item: any) => ({
        key: item.key,
        text: item.text
    }));

    return (
        <Panel
            isOpen={isOpen}
            onDismiss={onDismiss}
            type={PanelType.medium}
            headerText={isEditMode ? "Edit Folder" : "Add Folder"}
            isFooterAtBottom
            onRenderFooterContent={() => (
                <>
                    <PrimaryButton
                        text={isEditMode ? "Update" : "Save"}
                        onClick={handleSave}
                        styles={getPrimaryActionButtonStyles(8)}
                    />
                    <DefaultButton
                        text="Cancel"
                        onClick={onDismiss}
                        styles={getSecondaryActionButtonStyles()}
                    />
                </>
            )}
        >
            <Field>
                <label className="Headerlabel">
                    Template <span style={{ color: "red" }}>*</span>
                </label>
                <TextField
                    value={templateName}
                    disabled
                    readOnly
                    styles={{
                        field: {
                            background: "#f3f4f6",
                            border: "1px solid #e5e7eb",
                            borderRadius: 6,
                            color: "#374151",
                            fontSize: 14
                        }
                    }}
                />
            </Field>

            <Field>
                <label className="Headerlabel">
                    Folder Name <span style={{ color: "red" }}>*</span>
                </label>
                <TextField
                    value={folderName}
                    onChange={(_, val) => {
                        setFolderName(val || "");
                        setNameError("");
                    }}
                    placeholder="Enter Folder Name"
                    errorMessage={nameError}
                />
            </Field>

            <Field>
                <label className="Headerlabel">Is Child Folder?</label>
                <Toggle
                    checked={isChildFolder}
                    onChange={(_, val) => {
                        setIsChildFolder(!!val);
                        if (!val) setParentFolderId(undefined);
                        setParentError("");
                    }}
                />
            </Field>

            {isChildFolder && (
                <Field>
                    <label className="Headerlabel">
                        Parent Folder <span style={{ color: "red" }}>*</span>
                    </label>
                    <Select
                        options={parentFolderSelectOptions}
                        value={parentFolderSelectOptions.find((opt) => opt.key === parentFolderId)}
                        onChange={(selected: any) => {
                            setParentFolderId(selected?.key);
                            setParentError("");
                        }}
                        placeholder="Select Parent Folder"
                        getOptionLabel={(e: any) => e.text}
                        getOptionValue={(e: any) => String(e.key)}
                    />
                    {parentError && (
                        <p style={{ color: "red", fontSize: 12, marginTop: 5 }}>
                            {parentError}
                        </p>
                    )}
                </Field>
            )}

            <Field>
                <label className="Headerlabel">Active</label>
                <Toggle
                    checked={active}
                    onChange={(_, val) => setActive(!!val)}
                />
            </Field>
        </Panel>
    );
};

export default AddFolderPanel;
