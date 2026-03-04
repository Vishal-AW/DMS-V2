import * as React from "react";
import {
  DocumentFolder24Regular,
  Document24Regular,
  Archive24Regular,
  People24Regular,
  ChevronRight20Regular,
  Person20Regular,
  Calendar20Regular
} from '@fluentui/react-icons';
import { format } from "date-fns";

export interface Workspace {
  icon?: 'folder' | 'document' | 'archive' | 'team';
  accentColor?: string;
  Author: { Title: string; };
  Created?: string;
  ID: number;
  TileName: string;
  LibraryName: string;
  Order0: number;
  TileType?: string;
}

interface WorkspaceCardProps {
  workspace: Workspace;
  onClick: (workspace: Workspace) => void;
}

const iconMap = {
  folder: DocumentFolder24Regular,
  document: Document24Regular,
  archive: Archive24Regular,
  team: People24Regular,
};

export default function WorkspaceCard({ workspace, onClick }: WorkspaceCardProps) {
  const IconComponent = iconMap[workspace.icon || 'folder'];
  const accentColor = workspace.accentColor || '#0078d4';

  const cardStyle = { '--card-accent': accentColor } as React.CSSProperties;

  return (
    <div
      className="workspace-card"
      style={cardStyle}
      onClick={() => onClick(workspace)}
      data-testid={`card-workspace-${workspace.ID}`}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          onClick(workspace);
        }
      }}
      aria-label={`Open ${workspace.TileName} workspace`}
    >
      <div className="workspace-card-accent" />
      <div className="workspace-card-body">
        <div className="workspace-card-header-row">
          <div className="workspace-card-icon-wrap">
            <IconComponent className="workspace-card-icon-svg" />
          </div>
          <h3 className="workspace-card-title" data-testid={`text-workspace-title-${workspace.ID}`}>
            {workspace.TileName}
          </h3>
          <ChevronRight20Regular className="workspace-card-arrow" />
        </div>

        <div className="workspace-card-count" data-testid={`badge-doc-count-${workspace.ID}`}>
          {/* <span className="workspace-card-count-number">{workspace.documentCount}</span>
          <span className="workspace-card-count-label">Documents</span> */}
        </div>

        <div className="workspace-card-meta">
          {workspace.Author.Title && (
            <div className="workspace-card-meta-item" data-testid={`text-workspace-owner-${workspace.ID}`}>
              <Person20Regular className="workspace-card-meta-icon" />
              <span>{workspace.Author.Title}</span>
            </div>
          )}
          {workspace.Created && (
            <div className="workspace-card-meta-item" data-testid={`text-workspace-date-${workspace.ID}`}>
              <Calendar20Regular className="workspace-card-meta-icon" />
              <span>{format(workspace.Created, "MMM dd, yyyy")}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
