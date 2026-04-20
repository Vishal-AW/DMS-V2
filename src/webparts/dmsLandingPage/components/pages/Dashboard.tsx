/* eslint-disable */
import { SearchBox } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import WorkspaceCard, { Workspace } from "../../common/component/WorkspaceCard";
import { useNavigate } from "react-router-dom";
import { SPHttpClient } from "@microsoft/sp-http-base";
import SkeletonWidgets from "../../common/component/SkeletonWidgets";
import { getTileAccessGroupName } from "../../../../Services/TileService";

interface IDashboardProps {
    context: WebPartContext;
}

const Dashboard: React.FunctionComponent<IDashboardProps> = ({ context }) => {

    const navigate = useNavigate();
    const [searchQuery, setSearchQuery] = useState('');
    const [tileData, setTileData] = useState<Workspace[]>([]);
    const [userRole, setUserRole] = useState<string>("");
    const [groupIds, setGroupIds] = useState<number[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [userGroups, setUserGroups] = useState<any[]>([]);

    const SITEURL = context.pageContext.web.absoluteUrl;
    const USERID = context.pageContext.legacyPageContext.userId;

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

        setUserGroups(groups); // ✅ ADD THIS
        setGroupIds(groups.map((g: any) => g.Id));

        const isAdmin = groups.some((g: any) => g.Title === "ProjectAdmin");
        const isMember = groups.some((g: any) => g.Title === "Project Member");

        if (isAdmin) setUserRole("ProjectAdmin");
        else if (isMember) setUserRole("ProjectMember");
        else setUserRole("Guest");

    }, []);

    const hasTilePermission = (tile: any) => {

        if (userRole === "ProjectAdmin") return true;

        if (tile.TileAdminId === USERID) return true;


        const expectedGroupName = getTileAccessGroupName(tile.LibraryName);

        const hasGroupAccess = userGroups.some(
            (g: any) => g.Title === expectedGroupName
        );

        return hasGroupAccess;
    };

    const getTiles = useCallback(async () => {

        if (!userRole) return;

        setLoading(true);

        const query = `${SITEURL}/_api/web/lists/getByTitle('DMS_Mas_Tile')/items?$select=*,ID,TileName,Permission/ID,Permission/Title,TileAdmin/ID,Order0,icon,accentColor,Author/Title,LibraryName&$expand=Permission,TileAdmin,Author&$filter=Active eq 1&$orderby=Order0 asc`;

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


        const permittedTiles = tiles.filter(tile => hasTilePermission(tile));

        const permissionChecks = await Promise.all(
            permittedTiles.map(async (tile) => {
                const hasPermission = await checkLibraryPermission(tile.LibraryName);
                return hasPermission ? tile : null;
            })
        );

        const filteredTiles = permissionChecks
            .filter(Boolean)
            .sort((a: any, b: any) => Number(a.Order0) - Number(b.Order0));

        setTileData(filteredTiles as Workspace[]);
        setLoading(false);

    }, [userRole, groupIds]);


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
    }, [userRole, groupIds]);


    const filteredWorkspaces = useMemo(() => {
        if (!searchQuery.trim()) return tileData;

        const query = searchQuery.toLowerCase();

        return tileData.filter(ws =>
            Object.keys(ws).some(key => {
                const value = (ws as any)[key];

                return (
                    value !== null &&
                    value !== undefined &&
                    value.toString().toLowerCase().includes(query)
                );
            })
        );

    }, [searchQuery, tileData]);

    // ✅ LOADING UI
    if (loading) return (
        <div className="workspace-grid">
            {Array.from({ length: 6 }).map((_, index) => (
                <SkeletonWidgets key={index} />
            ))}
        </div>
    );

    // ✅ FINAL UI
    return (
        <div className="content-area-full" data-testid="page-dashboard">

            <div className="dashboard-header">
                <h2 className="dashboard-title">Department Workspaces</h2>
            </div>

            <div className="dashboard-search">
                <SearchBox
                    placeholder="Search workspaces..."
                    value={searchQuery}
                    onChange={(_, value) => setSearchQuery(value || '')}
                    onClear={() => setSearchQuery('')}
                    className="dashboard-search-box"
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
                    <div className="workspace-no-results">
                        No workspaces match your search.
                    </div>
                )}
            </div>

        </div>
    );
};

export default React.memo(Dashboard);