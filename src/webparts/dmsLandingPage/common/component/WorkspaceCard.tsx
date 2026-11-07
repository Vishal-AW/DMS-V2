import * as React from "react";
import {
  DocumentFolder24Regular,
  Document24Regular,
  Archive24Regular,
  People24Regular,
  ChevronRight20Regular
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

  // Initials for the owner avatar (e.g. "Chris Roy" -> "CR")
  const ownerName = workspace.Author?.Title || '';
  const ownerInitials = ownerName
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part.charAt(0))
    .join('')
    .toUpperCase();

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
      <span className="workspace-card-glow" aria-hidden="true" />
      <div className="workspace-card-body">
        <div className="workspace-card-header-row">
          <div className="workspace-card-icon-wrap">
            <IconComponent className="workspace-card-icon-svg" />
          </div>
          <h3 className="workspace-card-title" data-testid={`text-workspace-title-${workspace.ID}`}>
            {workspace.TileName}
          </h3>
        </div>

        <div className="workspace-card-count" data-testid={`badge-doc-count-${workspace.ID}`}>
          {/* <span className="workspace-card-count-number">{workspace.documentCount}</span>
          <span className="workspace-card-count-label">Documents</span> */}
        </div>

        <div className="workspace-card-footer">
          {ownerName && (
            <div className="workspace-card-owner">
              <span className="workspace-card-avatar" aria-hidden="true">
                {ownerInitials}
              </span>
              <span className="workspace-card-owner-text">
                <span
                  className="workspace-card-owner-name"
                  data-testid={`text-workspace-owner-${workspace.ID}`}
                >
                  {ownerName}
                </span>
                {workspace.Created && (
                  <span className="workspace-card-owner-date">
                    Created{' '}
                    <span data-testid={`text-workspace-date-${workspace.ID}`}>
                      {format(workspace.Created, "MMM dd, yyyy")}
                    </span>
                  </span>
                )}
              </span>
            </div>
          )}
          <span className="workspace-card-chevron" aria-hidden="true">
            <ChevronRight20Regular className="workspace-card-arrow" />
          </span>
        </div>
      </div>
    </div>
  );
}
