import * as React from "react";
import { IconButton, Text } from '@fluentui/react';
import {
  Delete20Regular,
  CheckmarkCircle20Regular,
  Search20Regular,
  Send20Regular
} from '@fluentui/react-icons';
import FolderTree, { FolderNode } from './FolderTree';
import { IDmsButton } from "../../components/pages/buttonPermissionHelper";

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
  /**
   * TODO (Transmittals): there is no Transmittals route, handler or count
   * source in the app yet. These two optional props let the item behave
   * exactly like "Approval" as soon as the caller supplies them; until then
   * the pill renders 0 and the click is a no-op.
   */
  transmittalsCount?: number;
  onTransmittalsClick?: () => void;
  LibDetails: any;
  buttons: any;
  permittedButtons?: IDmsButton[];
  expandedFolders: any;
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
  transmittalsCount = 0,
  onTransmittalsClick,
  LibDetails,
  archiveCount,
  buttons,
  permittedButtons = [],
  expandedFolders
}: SidebarProps) => {
  const isButtonPermitted = (internalName: string) => {
    // return permittedButtons.some(btn => btn.InternalName === internalName);
     return true;
  };

  return (
    <div className="sidebar1" data-testid="container-sidebar">
      <div className="sidebar-quick-links">

        {isButtonPermitted("Archive") && LibDetails?.IsArchiveRequired && (
          <div
            className="sidebar-quick-link"
            onClick={onArchiveClick}
            data-testid="link-recycle-bin"
            role="button"
            tabIndex={0}
          >
            <Delete20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-red" />
            <span className="sidebar-quick-link-label">Archive</span>
            <span className="sidebar-quick-link-count">{archiveCount ?? 0}</span>
          </div>
        )}

        {isButtonPermitted("RecycleBin") && !LibDetails?.IsArchiveRequired && (
          <div
            className="sidebar-quick-link"
            onClick={onRecycleBinClick}
            data-testid="link-recycle-bin"
            role="button"
            tabIndex={0}
          >
            <Delete20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-red" />
            <span className="sidebar-quick-link-label">Recycle Bin</span>
            <span className="sidebar-quick-link-count">{recycleBinCount}</span>
          </div>
        )}

        {isButtonPermitted("Approval") && (
          <div
            className="sidebar-quick-link"
            onClick={onApprovalClick}
            data-testid="link-approval"
            role="button"
            tabIndex={0}
          >
            <CheckmarkCircle20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-green" />
            <span className="sidebar-quick-link-label">Approval</span>
            <span className="sidebar-quick-link-count">{approvalCount}</span>
          </div>
        )}

        {/* NEW: Transmittals - same pattern as Approval (icon + label + count pill).
            TODO: no route/handler/count source exists yet, so the pill is 0 and the
            click does nothing until onTransmittalsClick / transmittalsCount are passed in. */}
        {isButtonPermitted("Transmittals") && (
          <div
            className="sidebar-quick-link"
            onClick={onTransmittalsClick}
            data-testid="link-transmittals"
            role="button"
            tabIndex={0}
          >
            <Send20Regular className="sidebar-quick-link-icon sidebar-quick-link-icon-blue" />
            <span className="sidebar-quick-link-label">Transmittals</span>
            <span className="sidebar-quick-link-count">{transmittalsCount}</span>
          </div>
        )}

        {isButtonPermitted("AdvancedSearch") && (
          <div
            className="sidebar-quick-link"
            onClick={onAdvancedSearchClick}
            data-testid="link-advanced-search"
            role="button"
            tabIndex={0}
          >
            <Search20Regular className="sidebar-quick-link-icon" />
            <span className="sidebar-quick-link-label">Advanced Search</span>
          </div>
        )}
      </div>

      <div className="sidebar-divider" />

      <div className="sidebar-content">
        <FolderTree
          folders={folders}
          selectedId={selectedFolderId}
          onFolderSelect={onFolderSelect}
          onFolderAction={onFolderAction}
          buttons={buttons}
          expandedFolders={expandedFolders}
        />
      </div>
    </div>
  );
};

export default React.memo(Sidebar);