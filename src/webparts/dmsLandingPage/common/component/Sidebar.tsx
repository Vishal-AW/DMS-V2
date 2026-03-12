import * as React from "react";
import { IconButton, Text } from '@fluentui/react';
import {
  Delete20Regular,
  CheckmarkCircle20Regular,
  Search20Regular
} from '@fluentui/react-icons';
import FolderTree, { FolderNode } from './FolderTree';

interface SidebarProps {
  folders: FolderNode[];
  selectedFolderId?: string;
  onFolderSelect: (folder: FolderNode) => void;
  onFolderAction?: (action: string, folder: FolderNode) => void;
  title?: string;
  recycleBinCount?: number;
  approvalCount?: number;
  archiveCount?: number;
  onRecycleBinClick?: () => void;
  onArchiveClick?: () => void;
  onApprovalClick?: () => void;
  onAdvancedSearchClick?: () => void;
  LibDetails: any;
  buttons: any;
}

const Sidebar = ({
  folders,
  selectedFolderId,
  onFolderSelect,
  onFolderAction,
  title = 'Folders',
  recycleBinCount = 0,
  approvalCount = 0,
  onRecycleBinClick,
  onApprovalClick,
  onAdvancedSearchClick,
  onArchiveClick,
  LibDetails,
  archiveCount,
  buttons
}: SidebarProps) => {
  return (
    <div className="sidebar1" data-testid="container-sidebar">
      <div className="sidebar-quick-links">

        {
          LibDetails?.IsArchiveRequired ?
            <div
              className="sidebar-quick-link"
              onClick={onArchiveClick}
              data-testid="link-recycle-bin"
              role="button"
              tabIndex={0}
            >
              <Delete20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-red" />
              <span>Archive({archiveCount})</span>
            </div>
            :
            <div
              className="sidebar-quick-link"
              onClick={onRecycleBinClick}
              data-testid="link-recycle-bin"
              role="button"
              tabIndex={0}
            >
              <Delete20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-red" />
              <span>Recycle Bin ({recycleBinCount})</span>
            </div>
        }



        <div
          className="sidebar-quick-link"
          onClick={onApprovalClick}
          data-testid="link-approval"
          role="button"
          tabIndex={0}
        >
          <CheckmarkCircle20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-green" />
          <span>Approval ({approvalCount})</span>
        </div>
        <div
          className="sidebar-quick-link"
          onClick={onAdvancedSearchClick}
          data-testid="link-advanced-search"
          role="button"
          tabIndex={0}
        >
          <Search20Regular className="sidebar-quick-link-icon" />
          <span>Advanced Search</span>
        </div>
      </div>

      <div className="sidebar-divider" />

      <div className="sidebar-content">
        <FolderTree
          folders={folders}
          selectedId={selectedFolderId}
          onFolderSelect={onFolderSelect}
          onFolderAction={onFolderAction}
          buttons={buttons}
        />
      </div>
    </div>
  );
};

export default React.memo(Sidebar);