/* eslint-disable */
import * as React from "react";
import { useState, useCallback, useRef, useEffect } from "react";
import {
    DndContext,
    DragOverlay,
    closestCenter,
    PointerSensor,
    useSensor,
    useSensors,
    DragStartEvent,
    DragEndEvent,
    DragOverEvent,
    UniqueIdentifier
} from "@dnd-kit/core";
import {
    SortableContext,
    verticalListSortingStrategy,
    useSortable,
    arrayMove
} from "@dnd-kit/sortable";
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

// Flatten tree to a list of IDs for dnd-kit sorting
const flattenTree = (nodes: FolderNode[]): string[] => {
    const ids: string[] = [];
    for (const node of nodes) {
        ids.push(`folder-${node.ID}`);
        if (node.children && node.children.length > 0 && node.isExpanded !== false) {
            ids.push(...flattenTree(node.children));
        }
    }
    return ids;
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
    onDragEnd: (draggedFolderId: number, newParentId: number | null, newOrder: number) => void;
    isLoading?: boolean;
}

interface SortableFolderItemProps {
    node: FolderNode;
    depth: number;
    isLast: boolean;
    parentLines: boolean[];
    allFolders: FolderNode[];
    onToggleExpand: (id: number) => void;
    onDragEnd: (draggedFolderId: number, newParentId: number | null, newOrder: number) => void;
    onRefreshTree: () => void;
    isDragOverlay?: boolean;
}

const SortableFolderItem = ({
    node,
    depth,
    isLast,
    parentLines,
    allFolders,
    onToggleExpand,
    onDragEnd,
    onRefreshTree,
    isDragOverlay = false
}: SortableFolderItemProps): JSX.Element => {
    const {
        attributes,
        listeners,
        setNodeRef,
        transform,
        transition,
        isDragging
    } = useSortable({
        id: `folder-${node.ID}`,
        data: {
            type: 'folder',
            node
        }
    });

    const hasChildren = node.children && node.children.length > 0;
    const INDENT = 24;
    const LINE_X = 16;

    const style: React.CSSProperties = {
        transform: CSS.Transform.toString(transform),
        transition,
        opacity: isDragging ? 0.4 : 1,
        position: "relative" as const,
        zIndex: isDragging ? 1 : "auto"
    };

    return (
        <div ref={setNodeRef} style={style}>
            <div
                className="folder-tree-row"
                style={{
                    position: "relative",
                    display: "flex",
                    alignItems: "center",
                    minHeight: 36,
                    fontSize: 14,
                    cursor: "default",
                    userSelect: "none",
                    borderRadius: 4,
                    background: isDragOverlay ? "#e8f4fd" : "transparent",
                    border: isDragOverlay ? "1px solid #0078d4" : "none",
                    boxShadow: isDragOverlay ? "0 4px 12px rgba(0,0,0,0.15)" : "none"
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
                        title="Drag to reorder"
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
                </div>
            </div>

            {/* Children */}
            {hasChildren && node.isExpanded !== false && (
                <SortableContext
                    items={node.children!.map(c => `folder-${c.ID}`)}
                    strategy={verticalListSortingStrategy}
                >
                    {node.children!.map((child: FolderNode, index: number) => (
                        <SortableFolderItem
                            key={child.ID}
                            node={child}
                            depth={depth + 1}
                            isLast={index === node.children!.length - 1}
                            parentLines={[...parentLines, !isLast]}
                            allFolders={allFolders}
                            onToggleExpand={onToggleExpand}
                            onDragEnd={onDragEnd}
                            onRefreshTree={onRefreshTree}
                        />
                    ))}
                </SortableContext>
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
    onDragEnd,
    isLoading = false
}) => {
    const [treeData, setTreeData] = useState<FolderNode[]>(folders);
    const [activeId, setActiveId] = useState<UniqueIdentifier | null>(null);
    const [overId, setOverId] = useState<UniqueIdentifier | null>(null);
    const [dropPosition, setDropPosition] = useState<"above" | "below" | "inside" | null>(null);
    const autoExpandTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

    // Sync treeData when folders prop changes
    useEffect(() => {
        setTreeData(folders);
    }, [folders]);

    const sensors = useSensors(
        useSensor(PointerSensor, {
            activationConstraint: {
                distance: 5 // 5px movement required to start drag
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
            setDropPosition(null);
            return;
        }

        setOverId(over.id);

        // Auto-expand logic: if hovering over a collapsed folder, expand after delay
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
        setDropPosition(null);

        if (autoExpandTimerRef.current) {
            clearTimeout(autoExpandTimerRef.current);
            autoExpandTimerRef.current = null;
        }

        if (!over) return;

        const activeIdStr = String(active.id);
        const overIdStr = String(over.id);

        if (activeIdStr === overIdStr) return;

        const activeFolderId = parseInt(activeIdStr.replace('folder-', ''));
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

        // Determine new parent based on drop position
        // For dnd-kit with sortable, items are reordered within the same container
        // We need to determine if the drop is "inside" (making it a child) or "between" (reordering)
        // Since dnd-kit sortable handles reordering within the same level,
        // we need to detect if the user wants to make it a child

        // For now, we'll use a simple approach:
        // If the active item is dropped on a folder that has children, make it a child
        // Otherwise, reorder at the same level

        // The new parent is the over node's parent (for reordering) or the over node itself (for nesting)
        // We'll determine this based on the drop position

        // Default: reorder at the same level as the over node
        let newParentId: number | null = overNode.ParentFolderIdId;

        // If the over node has children or is expanded, we can drop inside
        // For simplicity, we'll use the over node's parent as the new parent
        // This means dropping on a folder reorders within that folder's parent

        // Remove the active node from its current position
        const { node: removedNode, newTree } = removeNodeFromTree(treeData, activeFolderId);
        if (!removedNode) return;

        // Insert the removed node at the new position
        // For reordering within the same parent, we use the over node's position
        const updatedNode = { ...removedNode, ParentFolderIdId: newParentId };
        const finalTree = insertNodeIntoTree(newTree, newParentId, updatedNode);

        setTreeData(finalTree);

        // Call the parent's onDragEnd to persist changes
        onDragEnd(activeFolderId, newParentId, 1);
    }, [treeData, onDragEnd]);

    const handleDragCancel = useCallback(() => {
        setActiveId(null);
        setOverId(null);
        setDropPosition(null);
        if (autoExpandTimerRef.current) {
            clearTimeout(autoExpandTimerRef.current);
            autoExpandTimerRef.current = null;
        }
    }, []);

    const activeNode = activeId ? findNodeById(treeData, parseInt(String(activeId).replace('folder-', ''))) : null;

    // Flatten tree for sortable context
    const flatIds = flattenTree(treeData);

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
            collisionDetection={closestCenter}
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
                <SortableContext items={flatIds} strategy={verticalListSortingStrategy}>
                    {treeData.map((node: FolderNode, index: number) => (
                        <SortableFolderItem
                            key={node.ID}
                            node={node}
                            depth={0}
                            isLast={index === treeData.length - 1}
                            parentLines={[]}
                            allFolders={treeData}
                            onToggleExpand={handleToggleExpand}
                            onDragEnd={onDragEnd}
                            onRefreshTree={onRefreshTree}
                        />
                    ))}
                </SortableContext>
            </div>

            {/* Drag Overlay */}
            <DragOverlay>
                {activeNode ? (
                    <SortableFolderItem
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
