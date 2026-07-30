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

    const getTemplateFolderCount = (templateId: number): number => {
        const tree = templateFolderTrees[templateId];
        if (!tree) return 0;
        const countFolders = (nodes: FolderNode[]): number => {
            let count = 0;
            for (const node of nodes) {
                count += 1;
                if (node.children) {
                    count += countFolders(node.children);
                }
            }
            return count;
        };
        return countFolders(tree);
    };

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
                const folderCount = getTemplateFolderCount(template.ID);
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

                            {/* Folder Count */}
                            {folderCount > 0 && (
                                <span
                                    style={{
                                        fontSize: 12,
                                        color: "#6b7280",
                                        background: "#f3f4f6",
                                        padding: "2px 8px",
                                        borderRadius: 4,
                                        marginRight: 8
                                    }}
                                >
                                    {folderCount} folder{folderCount !== 1 ? "s" : ""}
                                </span>
                            )}

                            {/* Edit Button */}
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
                                    background: "#f5f8fa",
                                    marginLeft: 4
                                }}
                                title="Edit Template"
                            >
                                <FontIcon iconName="EditSolid12" />
                            </span>
                        </div>

                        {/* Accordion Content */}
                        {isExpanded && (
                            <div style={{ padding: "16px 20px 20px" }}>
                                {/* Folder Tree */}
                                <FolderTreeView
                                    folders={tree}
                                    templateId={template.ID}
                                    templateName={template.Name}
                                    onRefreshTree={() => handleRefreshTree(template.ID)}
                                    onAddFolder={() => handleAddFolder(template.ID, template.Name)}
                                    onDragEnd={handleDragEnd}
                                    isLoading={isLoading}
                                />

                                {/* Add Folder Button */}
                                <div style={{ marginTop: 12 }}>
                                    <PrimaryButton
                                        text="+ Add Folder"
                                        onClick={() => handleAddFolder(template.ID, template.Name)}
                                        styles={getAddActionButtonStyles()}
                                    />
                                </div>
                            </div>
                        )}
                    </div>
                );
            })}

            {/* Add Folder Panel */}
            {selectedTemplateForAdd && (
                <AddFolderPanel
                    context={context}
                    isOpen={addFolderPanelOpen}
                    onDismiss={() => {
                        setAddFolderPanelOpen(false);
                        setSelectedTemplateForAdd(null);
                    }}
                    onSaved={handleAddFolderSaved}
                    templateId={selectedTemplateForAdd.id}
                    templateName={selectedTemplateForAdd.name}
                />
            )}
        </div>
    );
};

export default TemplateAccordion;
