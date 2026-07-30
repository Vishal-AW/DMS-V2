/* eslint-disable */
import * as React from "react";
import { useState, useCallback, useRef } from "react";
import { FontIcon, PrimaryButton } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import FolderTreeView, { buildFolderTree, FolderNode } from "./FolderTreeView";
import AddFolderPanel from "./AddFolderPanel";
import { getFoldersByTemplateId, UpdateFolderMaster } from "../../../../Services/FolderMasterService";
import { getAddActionButtonStyles } from "../../common/component/buttonStyles";

interface TemplateAccordionProps {
    context: WebPartContext;
    templates: any[];
    onRefreshTemplates: () => void;
    onEditTemplate?: (id: number) => void;
}

const TemplateAccordion: React.FC<TemplateAccordionProps> = ({
    context,
    templates,
    onRefreshTemplates,
    onEditTemplate
}) => {
    const [expandedTemplateId, setExpandedTemplateId] = useState<number | null>(null);
    const [templateFolderTrees, setTemplateFolderTrees] = useState<Record<number, FolderNode[]>>({});
    const [loadingTemplateId, setLoadingTemplateId] = useState<number | null>(null);
    const [addFolderPanelOpen, setAddFolderPanelOpen] = useState(false);
    const [selectedTemplateForAdd, setSelectedTemplateForAdd] = useState<{ id: number; name: string; } | null>(null);
    const [editFolderData, setEditFolderData] = useState<any | null>(null);
    const loadedTemplatesRef = useRef<Set<number>>(new Set());

    const toggleAccordion = useCallback(async (templateId: number) => {
        // Check if already expanded - collapse it
        setExpandedTemplateId(prev => {
            if (prev === templateId) {
                return null;
            }
            return templateId;
        });

        // If already loaded, no need to fetch again
        if (loadedTemplatesRef.current.has(templateId)) return;

        setLoadingTemplateId(templateId);
        try {
            const res: any = await getFoldersByTemplateId(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                templateId
            );
            const folders = res?.value || [];
            const tree = buildFolderTree(folders, null);
            loadedTemplatesRef.current.add(templateId);
            setTemplateFolderTrees(prev => ({
                ...prev,
                [templateId]: tree
            }));
        } catch (error) {
            console.error("Error loading folders for template:", error);
            loadedTemplatesRef.current.add(templateId);
            setTemplateFolderTrees(prev => ({
                ...prev,
                [templateId]: []
            }));
        } finally {
            setLoadingTemplateId(null);
        }
    }, [context]);

    const handleRefreshTree = useCallback(async (templateId: number) => {
        try {
            const res: any = await getFoldersByTemplateId(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                templateId
            );
            const folders = res?.value || [];
            const tree = buildFolderTree(folders, null);
            setTemplateFolderTrees(prev => ({
                ...prev,
                [templateId]: tree
            }));
        } catch (error) {
            console.error("Error refreshing folder tree:", error);
        }
    }, [context]);

    const handleAddFolder = useCallback((templateId: number, templateName: string) => {
        setEditFolderData(null);
        setSelectedTemplateForAdd({ id: templateId, name: templateName });
        setAddFolderPanelOpen(true);
    }, []);

    const handleEditFolder = useCallback((templateId: number, templateName: string, folderData: any) => {
        setEditFolderData(folderData);
        setSelectedTemplateForAdd({ id: templateId, name: templateName });
        setAddFolderPanelOpen(true);
    }, []);

    const handleAddFolderSaved = useCallback(() => {
        if (selectedTemplateForAdd) {
            handleRefreshTree(selectedTemplateForAdd.id);
        }
    }, [selectedTemplateForAdd, handleRefreshTree]);

    const handleDragEnd = useCallback(async (draggedFolderId: number, newParentId: number | null, newOrder: number) => {
        if (!expandedTemplateId) return;

        try {
            const option: any = {
                ParentFolderIdId: newParentId
            };

            await UpdateFolderMaster(
                context.pageContext.web.absoluteUrl,
                context.spHttpClient,
                option,
                draggedFolderId
            );

            // Refresh the tree for the current template
            await handleRefreshTree(expandedTemplateId);
        } catch (error) {
            console.error("Error updating folder after drag:", error);
        }
    }, [expandedTemplateId, context, handleRefreshTree]);

    if (!templates || templates.length === 0) {
        return (
            <div
                style={{
                    padding: "24px",
                    textAlign: "center",
                    color: "#6b7280",
                    fontSize: 14,
                    border: "1px dashed #d1d5db",
                    borderRadius: 8,
                    background: "#f9fafb"
                }}
            >
                No templates found. Click "Add Template" to create one.
            </div>
        );
    }

    return (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {templates.map((template: any) => {
                const isExpanded = expandedTemplateId === template.ID;
                const isLoading = loadingTemplateId === template.ID;
                const tree = templateFolderTrees[template.ID] || [];

                return (
                    <div
                        key={template.ID}
                        style={{
                            border: "1px solid #e5e7eb",
                            borderRadius: 8,
                            background: "#ffffff",
                            overflow: "hidden",
                            boxShadow: isExpanded
                                ? "0 2px 8px rgba(0,0,0,0.08)"
                                : "0 1px 3px rgba(0,0,0,0.04)",
                            transition: "box-shadow 0.2s ease"
                        }}
                    >
                        {/* Accordion Header */}
                        <div
                            style={{
                                display: "flex",
                                alignItems: "center",
                                padding: "14px 20px",
                                cursor: "pointer",
                                background: isExpanded ? "#f8faff" : "#ffffff",
                                borderBottom: isExpanded ? "1px solid #e5e7eb" : "none",
                                transition: "background 0.15s ease",
                                userSelect: "none"
                            }}
                            onClick={() => toggleAccordion(template.ID)}
                        >
                            {/* Chevron */}
                            <span
                                style={{
                                    display: "inline-flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    width: 20,
                                    height: 20,
                                    marginRight: 12,
                                    color: "#6b7280",
                                    transform: isExpanded ? "rotate(90deg)" : "rotate(0deg)",
                                    transition: "transform 0.2s ease",
                                    fontSize: 16
                                }}
                            >
                                <svg width="12" height="12" viewBox="0 0 12 12" fill="none">
                                    <path d="M4.5 2.5L8.5 6L4.5 9.5" stroke="#6b7280" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                                </svg>
                            </span>

                            {/* Template Icon */}
                            <span style={{ marginRight: 10, fontSize: 18 }}>
                                {isExpanded ? "📋" : "📄"}
                            </span>

                            {/* Template Name */}
                            <span
                                style={{
                                    fontWeight: 600,
                                    fontSize: 15,
                                    color: "#111827",
                                    flex: 1
                                }}
                            >
                                {template.Name}
                            </span>

                            {/* Active Badge */}
                            <span
                                style={{
                                    display: "inline-flex",
                                    alignItems: "center",
                                    gap: 4,
                                    fontSize: 12,
                                    padding: "2px 10px",
                                    borderRadius: 12,
                                    background: template.Active ? "#ecfdf5" : "#f3f4f6",
                                    color: template.Active ? "#059669" : "#6b7280",
                                    fontWeight: 500,
                                    marginRight: 12
                                }}
                            >
                                <span
                                    style={{
                                        width: 6,
                                        height: 6,
                                        borderRadius: "50%",
                                        background: template.Active ? "#059669" : "#9ca3af"
                                    }}
                                />
                                {template.Active ? "Active" : "Inactive"}
                            </span>

                            {/* Action Buttons */}
                            <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                {/* Add Folder Button */}
                                <span
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        if (!template.Active) return;
                                        handleAddFolder(template.ID, template.Name);
                                    }}
                                    style={{
                                        display: "inline-flex",
                                        alignItems: "center",
                                        justifyContent: "center",
                                        gap: 4,
                                        padding: "5px 10px",
                                        borderRadius: 6,
                                        cursor: template.Active ? "pointer" : "not-allowed",
                                        color: template.Active ? "#009ef7" : "#d1d5db",
                                        background: template.Active ? "#f5f8fa" : "#f9fafb",
                                        fontSize: 12,
                                        fontWeight: 500,
                                        opacity: template.Active ? 1 : 0.6
                                    }}
                                    title={template.Active ? "Add Folder" : "Template is inactive"}
                                >
                                    <FontIcon iconName="Add" style={{ fontSize: 12 }} />
                                    Add Folder
                                </span>

                                {/* Edit Template Button */}
                                <span
                                    onClick={(e) => {
                                        e.stopPropagation();
                                        if (onEditTemplate) {
                                            onEditTemplate(template.ID);
                                        }
                                    }}
                                    style={{
                                        display: "inline-flex",
                                        alignItems: "center",
                                        justifyContent: "center",
                                        width: 32,
                                        height: 32,
                                        borderRadius: 6,
                                        cursor: "pointer",
                                        color: "#009ef7",
                                        background: "#f5f8fa"
                                    }}
                                    title="Edit Template"
                                >
                                    <FontIcon iconName="EditSolid12" />
                                </span>
                            </div>
                        </div>

                        {/* Accordion Content */}
                        {isExpanded && (
                            <div style={{ padding: "16px 20px 20px" }}>
                                {/* Folder Tree */}
                                <FolderTreeView
                                    folders={tree}
                                    templateId={template.ID}
                                    templateName={template.Name}
                                    templateActive={template.Active}
                                    onRefreshTree={() => handleRefreshTree(template.ID)}
                                    onAddFolder={() => handleAddFolder(template.ID, template.Name)}
                                    onEditFolder={(folderData) => handleEditFolder(template.ID, template.Name, folderData)}
                                    onDragEnd={handleDragEnd}
                                    isLoading={isLoading}
                                />
                            </div>
                        )}
                    </div>
                );
            })}

            {/* Add/Edit Folder Panel */}
            {selectedTemplateForAdd && (
                <AddFolderPanel
                    context={context}
                    isOpen={addFolderPanelOpen}
                    onDismiss={() => {
                        setAddFolderPanelOpen(false);
                        setSelectedTemplateForAdd(null);
                        setEditFolderData(null);
                    }}
                    onSaved={handleAddFolderSaved}
                    templateId={selectedTemplateForAdd.id}
                    templateName={selectedTemplateForAdd.name}
                    editFolderData={editFolderData}
                />
            )}
        </div>
    );
};

export default TemplateAccordion;
