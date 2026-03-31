import * as React from "react";
import { TooltipHost } from '@fluentui/react';
import {
  Folder20Regular,
  Folder20Filled,
  FolderOpen20Filled,
  FolderArrowRight20Regular,
  ChevronRight24Regular,
  MoreHorizontal24Regular,
  MoreHorizontalRegular
} from '@fluentui/react-icons';
import { Button, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger } from "@fluentui/react-components";
import * as FluentIcons from "@fluentui/react-icons";

export interface FolderNode {
  id: string;
  name: string;
  children?: FolderNode[];
  isLastLevel?: boolean;
  path: string;
}

interface FolderTreeProps {
  folders: FolderNode[];
  selectedId?: string;
  onFolderSelect: (folder: FolderNode) => void;
  onFolderAction?: (action: string, folder: FolderNode) => void;
  buttons: any;
  expandedFolders: any;
}

interface FolderTreeItemProps {
  folder: FolderNode;
  level: number;
  selectedId?: string;
  onSelect: (folder: FolderNode) => void;
  onFolderAction?: (action: string, folder: FolderNode) => void;
  buttons: any;
  showButton: boolean;
  expandedFolders: any;
}

function FolderTreeItem({ folder, level, selectedId, onSelect, onFolderAction, buttons, showButton, expandedFolders }: FolderTreeItemProps) {

  const hasChildren = folder.children && folder.children.length > 0;
  const isSelected = folder.id === selectedId;
  const isLeaf = folder.isLastLevel || (!hasChildren);
  const isExpanded = expandedFolders.includes(folder.id);

  const handleClick = () => {
    onSelect(folder);
  };

  const leftIndent = 10 + Math.min(level, 6) * 14;



  const renderFolderIcon = () => {
    if (isLeaf) {
      return <Folder20Filled className="folder-tree-icon folder-tree-icon-leaf" />;
    }
    if (isExpanded) {
      return <FolderOpen20Filled className="folder-tree-icon folder-tree-icon-open" />;
    }
    if (hasChildren && level === 0) {
      return <FolderArrowRight20Regular className="folder-tree-icon folder-tree-icon-root" />;
    }
    return <Folder20Regular className="folder-tree-icon folder-tree-icon-parent" />;
  };

  return (
    <>
      <div
        className={`folder-tree-item ${isSelected ? 'folder-tree-item-selected' : ''}`}
        data-testid={`folder-item-${folder.id}`}
        role="treeitem"
        aria-selected={isSelected}
        aria-expanded={hasChildren ? isExpanded : undefined}
        tabIndex={0}
        onKeyDown={(e) => {
          if (e.key === 'Enter' || e.key === ' ') {
            handleClick();
          }
        }}
      >
        <div
          className="folder-tree-item-content"
          style={{ paddingLeft: `${leftIndent}px` }}
          onClick={handleClick}
        >
          {renderFolderIcon()}

          <span className="folder-tree-name" data-testid={`text-folder-name-${folder.id}`}>
            <TooltipHost content={folder.name}>
              {folder.name}
            </TooltipHost>
          </span>
          {onFolderAction && showButton && (
            <div className="folder-tree-actions-slot">
              <Menu>
                <MenuTrigger disableButtonEnhancement>
                  <Button
                    appearance="subtle"
                    className="folder-tree-actions"
                    icon={<MoreHorizontalRegular className="table-action-btn" />}
                  />
                </MenuTrigger>

                <MenuPopover
                  style={{
                    boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
                    padding: "15px"
                  }}
                >
                  <MenuList>
                    {buttons.map((e: any) => {
                      // const IconComponent = FluentIcons[e.Icons as keyof typeof FluentIcons] as React.FC ?? <ChevronRight24Regular />;
                      const IconComponent = (
                        FluentIcons[e.Icons as keyof typeof FluentIcons] ??
                        ChevronRight24Regular
                      ) as React.ComponentType<React.SVGProps<SVGSVGElement>>;
                      return <MenuItem
                        key={e?.key}
                        icon={<IconComponent className="table-action-btn" />}
                        onClick={() => onFolderAction(e?.key, folder)}
                      >
                        {e?.ButtonDisplayName}
                      </MenuItem>;
                    })}
                  </MenuList>
                </MenuPopover>
              </Menu>
            </div>
          )}
          {hasChildren && (
            <span className="folder-tree-count">{folder.children?.length}</span>
          )}
        </div>

      </div>
      {hasChildren && isExpanded && folder.children?.map((child) => (
        <FolderTreeItem
          key={child.id}
          folder={child}
          level={level + 1}
          selectedId={selectedId}
          onSelect={onSelect}
          onFolderAction={onFolderAction}
          buttons={buttons}
          showButton={true}
          expandedFolders={expandedFolders}
        />
      ))}
    </>
  );
}

export default function FolderTree({ folders, selectedId, onFolderSelect, onFolderAction, buttons, expandedFolders }: FolderTreeProps) {
  return (
    <div role="tree" aria-label="Folder navigation" data-testid="tree-folders">
      {folders.map((folder, index) => (
        <FolderTreeItem
          key={folder.id}
          folder={folder}
          level={0}
          selectedId={selectedId}
          onSelect={onFolderSelect}
          onFolderAction={onFolderAction}
          buttons={buttons}
          showButton={index !== 0}
          expandedFolders={expandedFolders}
        />
      ))}
    </div>
  );
}
