/* eslint-disable */
import * as React from "react";
import { useState, useCallback, useRef, useEffect } from "react";
import { FontIcon } from "@fluentui/react";
import {
    DndContext,
    DragOverlay,
    pointerWithin,
    PointerSensor,
    useSensor,
    useSensors,
    DragStartEvent,
    DragEndEvent,
    DragOverEvent,
    UniqueIdentifier,
    useDraggable,
    useDroppable
} from "@dnd-kit/core";
import { CSS } from "@dnd-kit/utilities";

// Interface for the nested folder structure
export interface FolderNode {
    ID: number;
    FolderName: string;
    IsParentFolder: boolean;
    Active: boolean;
    ParentFolderIdId: number | null;
    children?: FolderNode[];
    isExpanded?: boolean;
    TemplateNameId?: number;
}

// Helper function to build the folder tree from flat list
export const buildFolderTree = (folders: any[], parentId: number | null = null): FolderNode[] => {
    const tree: FolderNode[] = [];

    const directChildren = folders.filter(folder => {
        const folderParentId = folder.ParentFolderIdId !== null && folder.ParentFolderIdId !== undefined
            ? Number(folder.ParentFolderIdId)
            : folder.ParentFolderId?.ID !== null && folder.ParentFolderId?.ID !== undefined
                ? Number(folder.ParentFolderId.ID)
                : null;
        return folderParentId === parentId;
    });

    for (const folder of directChildren) {
        const folderParentId = folder.ParentFolderIdId !== null && folder.ParentFolderIdId !== undefined
            ? Number(folder.ParentFolderIdId)
            : folder.ParentFolderId?.ID !== null && folder.ParentFolderId?.ID !== undefined
                ? Number(folder.ParentFolderId.ID)
                : null;

        const node: FolderNode = {
            ID: folder.ID,
            FolderName: folder.FolderName,
            IsParentFolder: folder.IsParentFolder,
            Active: folder.Active,
            ParentFolderIdId: folderParentId,
            TemplateNameId: folder.TemplateName?.ID,
            children: buildFolderTree(folders, folder.ID),
            isExpanded: true
        };
        tree.push(node);
    }
    return tree;
};

// Find a node in the tree by ID
const findNodeById = (nodes: FolderNode[], id: number): FolderNode | null => {
    for (const node of nodes) {
        if (node.ID === id) return node;
        if (node.children) {
            const found = findNodeById(node.children, id);
            if (found) return found;
        }
    }
    return null;
};

// Get all descendant IDs (including self)
const getDescendantIds = (node: FolderNode): number[] => {
    let ids: number[] = [node.ID];
    if (node.children) {
        for (const child of node.children) {
            ids = ids.concat(getDescendantIds(child));
        }
    }
    return ids;
};

// Remove a node from the tree by ID, returns the removed node and the new tree
const removeNodeFromTree = (nodes: FolderNode[], id: number): { node: FolderNode | null; newTree: FolderNode[]; } => {
    for (let i = 0; i < nodes.length; i++) {
        if (nodes[i].ID === id) {
            const removed = nodes[i];
            const newTree = [...nodes];
            newTree.splice(i, 1);
            return { node: removed, newTree };
        }
        if (nodes[i].children) {
            const result = removeNodeFromTree(nodes[i].children!, id);
            if (result.node) {
                const newChildren = [...nodes];
                newChildren[i] = { ...nodes[i], children: result.newTree };
                return { node: result.node, newTree: newChildren };
            }
        }
    }
    return { node: null, newTree: nodes };
};

// Insert a node into the tree at a specific parent
const insertNodeIntoTree = (
    nodes: FolderNode[],
    parentId: number | null,
    node: FolderNode,
    index: number = -1
): FolderNode[] => {
    if (parentId === null) {
        // Insert at root level
        const newTree = [...nodes];
        if (index >= 0 && index < newTree.length) {
            newTree.splice(index, 0, node);
        } else {
            newTree.push(node);
        }
        return newTree;
    }

    return nodes.map(n => {
        if (n.ID === parentId) {
            const newChildren = n.children ? [...n.children] : [];
            if (index >= 0 && index < newChildren.length) {
                newChildren.splice(index, 0, node);
            } else {
                newChildren.push(node);
            }
            return { ...n, children: newChildren, isExpanded: true };
        }
        if (n.children) {
            return { ...n, children: insertNodeIntoTree(n.children, parentId, node, index) };
        }
        return n;
    });
};

interface FolderTreeViewProps {
    folders: FolderNode[];
    templateId: number;
    templateName: string;
    onRefreshTree: () => void;
    onAddFolder: () => void;
    onEditFolder?: (folderData: any) => void;
    onDragEnd: (draggedFolderId: number, newParentId: number | null, newOrder: number) => void;
    isLoading?: boolean;
}

interface DraggableFolderItemProps {
    node: FolderNode;
    depth: number;
    isLast: boolean;
    parentLines: boolean[];
    allFolders: FolderNode[];
    onToggleExpand: (id: number) => void;
    onDragEnd: (draggedFolderId: number, newParentId: number | null, newOrder: number) => void;
    onRefreshTree: () => void;
    onEditFolder?: (folderData: any) => void;
    isDragOverlay?: boolean;
    isOver?: boolean;
}

// Droppable zone for each folder (to accept children)
const FolderDroppable = ({ id, children, isOver }: { id: string; children: React.ReactNode; isOver?: boolean; }) => {
    const { setNodeRef, isOver: isOverDroppable } = useDroppable({ id });
    return (
        <div
            ref={setNodeRef}
            style={{
                background: isOver || isOverDroppable ? "#e8f4fd" : "transparent",
                borderRadius: 4,
                transition: "background 0.15s ease"
            }}
        >
            {children}
        </div>
    );
};

// Root drop zone component - defined outside to avoid recreation on each render
const RootDroppable = () => {
    const { setNodeRef, isOver } = useDroppable({ id: "root-zone" });
    return (
        <div
            ref={setNodeRef}
            style={{
                minHeight: 40,
                border: isOver ? "2px dashed #0078d4" : "1px dashed transparent",
                borderRadius: 6,
                margin: "4px 8px",
                background: isOver ? "#f0f7ff" : "transparent",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                transition: "all 0.15s ease"
            }}
        >
            {isOver && (
                <span style={{ fontSize: 13, color: "#0078d4", fontWeight: 500 }}>
                    Drop here to make it a root folder
                </span>
            )}
        </div>
    );
};

const DraggableFolderItem = ({
    node,
    depth,
    isLast,
    parentLines,
    allFolders,
    onToggleExpand,
    onDragEnd,
    onRefreshTree,
    onEditFolder,
    isDragOverlay = false,
    isOver = false
}: DraggableFolderItemProps): JSX.Element => {
    const {
        attributes,
        listeners,
        setNodeRef: setDraggableRef,
        transform,
        isDragging
    } = useDraggable({
        id: `folder-${node.ID}`,
        data: {
            type: 'folder',
            node
        }
    });

    const hasChildren = node.children && node.children.length > 0;
    const isParentFolder = depth === 0;
    const INDENT = 24;
    const LINE_X = 16;

    const style: React.CSSProperties = {
        transform: CSS.Translate.toString(transform),
        opacity: isDragging ? 0.4 : 1,
        position: "relative" as const,
        zIndex: isDragging ? 1 : "auto"
    };

    return (
        <div ref={setDraggableRef} style={style}>
            <div
                className="folder-tree-row"
                style={{
                    position: "relative",
                    display: "flex",
                    alignItems: "center",
                    minHeight: 40,
                    fontSize: 14,
                    cursor: "default",
                    userSelect: "none",
                    borderRadius: 6,
                    padding: "6px 8px",
                    margin: "2px 6px",
                    background: isDragOverlay ? "#e8f4fd" : isOver ? "#f0f7ff" : "transparent",
                    border: isDragOverlay ? "1px solid #0078d4" : "none",
                    boxShadow: "none"
                }}
            >
                {/* Ancestor vertical lines */}
                {parentLines.map((show: boolean, idx: number) =>
                    show ? (
                        <div
                            key={idx}
                            style={{
                                position: "absolute",
                                left: idx * INDENT + LINE_X,
                                top: 0,
                                bottom: 0,
                                width: 1,
                                background: "#c0c0c0"
                            }}
                        />
                    ) : null
                )}

                {/* Current node: vertical + horizontal connector */}
                {depth > 0 && (
                    <>
                        <div
                            style={{
                                position: "absolute",
                                left: (depth - 1) * INDENT + LINE_X,
                                top: 0,
                                height: isLast ? "50%" : "100%",
                                width: 1,
                                background: "#c0c0c0"
                            }}
                        />
                        <div
                            style={{
                                position: "absolute",
                                left: (depth - 1) * INDENT + LINE_X,
                                top: "50%",
                                width: INDENT - LINE_X + 2,
                                height: 1,
                                background: "#c0c0c0"
                            }}
                        />
                    </>
                )}

                {/* Content row */}
                <div
                    style={{
                        display: "flex",
                        alignItems: "center",
                        paddingLeft: depth * INDENT + 8,
                        flex: 1,
                        gap: 4
                    }}
                >
                    {/* Chevron */}
                    <span
                        onClick={(e) => {
                            e.stopPropagation();
                            if (hasChildren) onToggleExpand(node.ID);
                        }}
                        style={{
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            width: 14,
                            marginRight: 2,
                            fontSize: 12,
                            color: "#6b7280",
                            visibility: hasChildren ? "visible" : "hidden",
                            transform: node.isExpanded ? "rotate(90deg)" : "rotate(0deg)",
                            transition: "transform 0.15s ease",
                            lineHeight: 1,
                            cursor: hasChildren ? "pointer" : "default"
                        }}
                    >
                        <svg width="8" height="12" viewBox="0 0 8 12" fill="none">
                            <path d="M1.5 1.5L6.5 6L1.5 10.5" stroke="#6b7280" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                        </svg>
                    </span>

                    {/* Drag handle */}
                    <span
                        {...attributes}
                        {...listeners}
                        style={{
                            display: "inline-flex",
                            alignItems: "center",
                            cursor: "grab",
                            color: "#9ca3af",
                            fontSize: 14,
                            marginRight: 2,
                            opacity: 0.5,
                            touchAction: "none"
                        }}
                        title="Drag to reorder or make child"
                    >
                        ⠿
                    </span>

                    {/* Folder icon */}
                    <span style={{ marginRight: 6, fontSize: 17, lineHeight: 1 }}>
                        {hasChildren && node.isExpanded ? "📂" : "📁"}
                    </span>

                    {/* Label */}
                    <span
                        style={{
                            fontWeight: hasChildren ? 600 : 400,
                            color: "#111827",
                            fontSize: 14,
                            flex: 1
                        }}
                    >
                        {node.FolderName}
                    </span>

                    {/* Active badge */}
                    {!node.Active && (
                        <span
                            style={{
                                fontSize: 11,
                                color: "#9ca3af",
                                background: "#f3f4f6",
                                padding: "1px 6px",
                                borderRadius: 4,
                                marginLeft: 8
                            }}
                        >
                            Inactive
                        </span>
                    )}

                    {/* Edit Folder Button */}
                    <span
                        onClick={(e) => {
                            e.stopPropagation();
                            if (onEditFolder) {
                                onEditFolder({
                                    ID: node.ID,
                                    FolderName: node.FolderName,
                                    Active: node.Active,
                                    ParentFolderIdId: node.ParentFolderIdId,
                                    TemplateNameId: node.TemplateNameId
                                });
                            }
                        }}
                        style={{
                            display: "inline-flex",
                            alignItems: "center",
                            justifyContent: "center",
                            cursor: "pointer",
                            color: "#009ef7",
                            backgroundColor: "#f5f8fa",
                            padding: "5px 8px",
                            borderRadius: "6px",
                            marginLeft: 6,
                            fontSize: 12
                        }}
                        title="Edit Folder"
                    >
                        <FontIcon iconName="EditSolid12" />
                    </span>
                </div>
            </div>

            {/* Children */}
            {hasChildren && node.isExpanded !== false && (
                <div style={{ position: "relative" }}>
                    {node.children!.map((child: FolderNode, index: number) => (
                        <FolderDroppable
                            key={child.ID}
                            id={`folder-${child.ID}`}
                            isOver={isOver}
                        >
                            <DraggableFolderItem
                                node={child}
                                depth={depth + 1}
                                isLast={index === node.children!.length - 1}
                                parentLines={[...parentLines, !isLast]}
                                allFolders={allFolders}
                                onToggleExpand={onToggleExpand}
                                onDragEnd={onDragEnd}
                                onRefreshTree={onRefreshTree}
                                onEditFolder={onEditFolder}
                                isOver={isOver}
                            />
                        </FolderDroppable>
                    ))}
                </div>
            )}
        </div>
    );
};

const FolderTreeView: React.FC<FolderTreeViewProps> = ({
    folders,
    templateId,
    templateName,
    onRefreshTree,
    onAddFolder,
    onEditFolder,
    onDragEnd,
    isLoading = false
}) => {
    const [treeData, setTreeData] = useState<FolderNode[]>(folders);
    const [activeId, setActiveId] = useState<UniqueIdentifier | null>(null);
    const [overId, setOverId] = useState<UniqueIdentifier | null>(null);
    const autoExpandTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

    // Sync treeData when folders prop changes
    useEffect(() => {
        setTreeData(folders);
    }, [folders]);

    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: {
                distance: 5
            }
        })
    );

    const handleToggleExpand = useCallback((id: number) => {
        setTreeData(prev => {
            const updateNode = (nodes: FolderNode[]): FolderNode[] => {
                return nodes.map(n => {
                    if (n.ID === id) {
                        return { ...n, isExpanded: !n.isExpanded };
                    }
                    if (n.children) {
                        return { ...n, children: updateNode(n.children) };
                    }
                    return n;
                });
            };
            return updateNode(prev);
        });
    }, []);

    const handleDragStart = useCallback((event: DragStartEvent) => {
        setActiveId(event.active.id);
    }, []);

    const handleDragOver = useCallback((event: DragOverEvent) => {
        const { over } = event;
        if (!over) {
            setOverId(null);
            return;
        }

        setOverId(over.id);

        // Auto-expand logic
        const overNode = findNodeById(treeData, parseInt(String(over.id).replace('folder-', '')));
        if (overNode && overNode.children && overNode.children.length > 0 && overNode.isExpanded === false) {
            if (!autoExpandTimerRef.current) {
                autoExpandTimerRef.current = setTimeout(() => {
                    handleToggleExpand(overNode.ID);
                    autoExpandTimerRef.current = null;
                }, 700);
            }
        } else {
            if (autoExpandTimerRef.current) {
                clearTimeout(autoExpandTimerRef.current);
                autoExpandTimerRef.current = null;
            }
        }
    }, [treeData, handleToggleExpand]);

    const handleDragEnd = useCallback((event: DragEndEvent) => {
        const { active, over } = event;
        setActiveId(null);
        setOverId(null);

        if (autoExpandTimerRef.current) {
            clearTimeout(autoExpandTimerRef.current);
            autoExpandTimerRef.current = null;
        }

        if (!over) return;

        const activeIdStr = String(active.id);
        const overIdStr = String(over.id);

        if (activeIdStr === overIdStr) return;

        const activeFolderId = parseInt(activeIdStr.replace('folder-', ''));

        // Check if dropped on root zone
        if (overIdStr === "root-zone") {
            // Make the active node a root folder
            const activeNode = findNodeById(treeData, activeFolderId);
            if (!activeNode) return;

            // Prevent dropping into any of its descendants (not applicable for root zone)
            const descendantIds = getDescendantIds(activeNode);
            if (descendantIds.includes(activeFolderId)) return;

            const newParentId: number | null = null;

            // Remove the active node from its current position
            const { node: removedNode, newTree } = removeNodeFromTree(treeData, activeFolderId);
            if (!removedNode) return;

            // Insert at root level
            const updatedNode = { ...removedNode, ParentFolderIdId: newParentId };
            const finalTree = insertNodeIntoTree(newTree, newParentId, updatedNode);
            setTreeData(finalTree);

            onDragEnd(activeFolderId, newParentId, 1);
            return;
        }

        const overFolderId = parseInt(overIdStr.replace('folder-', ''));

        // Get the active node from tree data
        const activeNode = findNodeById(treeData, activeFolderId);
        if (!activeNode) return;

        // Get the over node
        const overNode = findNodeById(treeData, overFolderId);
        if (!overNode) return;

        // Prevent dropping onto itself
        if (activeFolderId === overFolderId) return;

        // Prevent dropping into any of its descendants
        const descendantIds = getDescendantIds(activeNode);
        if (descendantIds.includes(overFolderId)) return;

        // Make the active node a child of the over node
        const newParentId: number | null = overFolderId;

        // Remove the active node from its current position
        const { node: removedNode, newTree } = removeNodeFromTree(treeData, activeFolderId);
        if (!removedNode) return;

        // Insert as a child of the over node
        const updatedNode = { ...removedNode, ParentFolderIdId: newParentId };
        const finalTree = insertNodeIntoTree(newTree, newParentId, updatedNode);
        setTreeData(finalTree);

        // Call the parent's onDragEnd to persist changes
        onDragEnd(activeFolderId, newParentId, 1);
    }, [treeData, onDragEnd]);

    const handleDragCancel = useCallback(() => {
        setActiveId(null);
        setOverId(null);
        if (autoExpandTimerRef.current) {
            clearTimeout(autoExpandTimerRef.current);
            autoExpandTimerRef.current = null;
        }
    }, []);

    const activeNode = activeId ? findNodeById(treeData, parseInt(String(activeId).replace('folder-', ''))) : null;

    if (isLoading) {
        return (
            <div style={{ padding: "16px 12px", color: "#6b7280", fontSize: 13 }}>
                Loading folder structure...
            </div>
        );
    }

    if (!treeData || treeData.length === 0) {
        return (
            <div
                style={{
                    padding: "16px 12px",
                    fontSize: 13,
                    color: "#6b7280",
                    textAlign: "center",
                    border: "1px dashed #d1d5db",
                    borderRadius: 6,
                    background: "#f9fafb"
                }}
            >
                No folders configured for this template.
            </div>
        );
    }

    return (
        <DndContext
            sensors={sensors}
            collisionDetection={pointerWithin}
            onDragStart={handleDragStart}
            onDragOver={handleDragOver}
            onDragEnd={handleDragEnd}
            onDragCancel={handleDragCancel}
        >
            <div
                style={{
                    border: "1px solid #e5e7eb",
                    borderRadius: 8,
                    background: "#ffffff",
                    padding: "4px 0",
                    marginTop: 8
                }}
            >
                {treeData.map((node: FolderNode, index: number) => (
                    <FolderDroppable
                        key={node.ID}
                        id={`folder-${node.ID}`}
                        isOver={overId === `folder-${node.ID}`}
                    >
                        <DraggableFolderItem
                            node={node}
                            depth={0}
                            isLast={index === treeData.length - 1}
                            parentLines={[]}
                            allFolders={treeData}
                            onToggleExpand={handleToggleExpand}
                            onDragEnd={onDragEnd}
                            onRefreshTree={onRefreshTree}
                            onEditFolder={onEditFolder}
                            isOver={overId === `folder-${node.ID}`}
                        />
                    </FolderDroppable>
                ))}

                {/* Root drop zone - drop folders here to make them root */}
                <RootDroppable />
            </div>

            {/* Drag Overlay */}
            <DragOverlay>
                {activeNode ? (
                    <DraggableFolderItem
                        node={activeNode}
                        depth={0}
                        isLast={false}
                        parentLines={[]}
                        allFolders={treeData}
                        onToggleExpand={() => { }}
                        onDragEnd={() => { }}
                        onRefreshTree={() => { }}
                        isDragOverlay={true}
                    />
                ) : null}
            </DragOverlay>
        </DndContext>
    );
};

export default FolderTreeView;
