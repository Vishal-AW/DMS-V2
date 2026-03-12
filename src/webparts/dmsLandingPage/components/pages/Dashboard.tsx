/* eslint-disable */
import { SearchBox } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import WorkspaceCard, { Workspace } from "../../common/component/WorkspaceCard";
import { useNavigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http-base";
import SkeletonWidgets from "../../common/component/SkeletonWidgets";

interface IDashboardProps {
    context: WebPartContext;
}


const Dashboard: React.FunctionComponent<IDashboardProps> = ({ context }) => {
    const navigate = useNavigate();
    const [searchQuery, setSearchQuery] = useState('');
    const [tileData, setTileData] = useState<Workspace[]>([]);
    const [userRole, setUserRole] = useState<string>("");
    const [groupIds, setGroupIds] = useState<number[]>([]);
    const SITEURL = context.pageContext.web.absoluteUrl;
    const USERID = context.pageContext.legacyPageContext.userId;
    const [loading, setLoading] = useState<boolean>(true);

    const handleWorkspaceClick = (workspace: Workspace) => {
        navigate(`/workspace/${workspace.ID}`);
    };

    const getUserRole = useCallback(async () => {
        const url = `${SITEURL}/_api/Web/GetUserById(${USERID})?$expand=Groups`;

        const response = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    Accept: "application/json;odata=nometadata",
                    "odata-version": ""
                }
            }
        );

        if (!response.ok) return;

        const data = await response.json();
        const groups = data.Groups || [];

        const isAdmin = groups.some((g: any) => g.Title === "ProjectAdmin");
        const isMember = groups.some((g: any) => g.Title === "Project Member");
        setGroupIds(data.Groups?.map((g: any) => g.Id) || []);
        if (isAdmin) setUserRole("ProjectAdmin");
        else if (isMember) setUserRole("ProjectMember");
        else setUserRole("Guest");
    }, []);

    const getTiles = useCallback(async () => {
        if (!userRole) return;

        setLoading(true);

        let filter = "Active eq 1";

        if (userRole !== "ProjectAdmin") {
            const groupFilter = groupIds
                .map(id => `Permission/ID eq ${id}`)
                .join(" or ");
            filter += ` and (Permission/ID eq ${USERID} or TileAdmin/ID eq ${USERID} ${groupFilter ? `or ${groupFilter}` : ""} or Permission/Title eq 'Everyone except external users')`;
        }

        const query = `${SITEURL}/_api/web/lists/getByTitle('DMS_Mas_Tile')/items?$select=*,ID,TileName,Permission/ID,Order0,icon,accentColor,Author/Title&$expand=Permission,Author&$filter=${filter}&$orderby=Order0`;

        const response = await context.spHttpClient.get(
            query,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    Accept: "application/json;odata=nometadata",
                    "odata-version": ""
                }
            }
        );

        if (!response.ok) {
            setLoading(false);
            return;
        }

        const data = await response.json();
        const tiles: Workspace[] = data.value || [];

        /* Check Library Permissions Properly */
        const permissionChecks = await Promise.all(
            tiles.map(async (tile) => {
                const hasPermission = await checkLibraryPermission(tile.LibraryName);
                return hasPermission ? tile : null;
            })
        );

        const filteredTiles = permissionChecks
            .filter(Boolean)
            .sort((a: any, b: any) => a.Order0 - b.Order0);

        setTileData(filteredTiles as Workspace[]);
        setLoading(false);
    }, [userRole]);

    const checkLibraryPermission = async (libraryName: string) => {
        try {
            const url = `${SITEURL}/_api/web/lists/GetByTitle('${libraryName}')/effectiveBasePermissions`;

            const response = await context.spHttpClient.get(
                url,
                SPHttpClient.configurations.v1
            );

            return response.ok;
        } catch {
            return false;
        }
    };

    useEffect(() => {
        getUserRole();
    }, []);

    useEffect(() => {
        if (userRole) getTiles();
    }, [userRole]);

    const filteredWorkspaces = useMemo(() => {
        if (!searchQuery.trim()) return tileData;

        const query = searchQuery.toLowerCase();

        return tileData.filter(ws =>
            Object.keys(ws).some(key => {
                const value = (ws as any)[key];
                return (
                    value !== null &&
                    value !== undefined &&
                    value
                        .toString()
                        .toLowerCase()
                        .includes(query)
                );
            })
        );
    }, [searchQuery, tileData]);

    if (loading) return <div className="workspace-grid">
        {
            Array.from({ length: 6 }).map((_, index) => (
                <SkeletonWidgets />
            ))
        }
    </div>;


    return (
        <div className="content-area-full" data-testid="page-dashboard">
            <div className="dashboard-header">
                <h2 className="dashboard-title" data-testid="text-page-title">Department Workspaces</h2>
            </div>
            <div className="dashboard-search">
                <SearchBox
                    placeholder="Search workspaces..."
                    value={searchQuery}
                    onChange={(_, value) => setSearchQuery(value || '')}
                    onClear={() => setSearchQuery('')}
                    className="dashboard-search-box"
                    data-testid="input-search-workspaces"
                />
            </div>
            <div className="workspace-grid">
                {filteredWorkspaces.length > 0 ? (
                    filteredWorkspaces.map(workspace => (
                        <WorkspaceCard
                            key={workspace.ID}
                            workspace={workspace}
                            onClick={handleWorkspaceClick}
                        />
                    ))
                ) : (
                    <div className="workspace-no-results" data-testid="text-no-workspaces">
                        No workspaces match your search.
                    </div>
                )}
            </div>
        </div>
    );
};

export default React.memo(Dashboard);