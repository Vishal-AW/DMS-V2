import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/site-users/web";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/security/web";
import { IPersonaProps } from "@fluentui/react";

export type TilePermissionLevel = "Read" | "Contribute";

export interface ITileAccessPrincipal extends IPersonaProps {
  principalId: number;
  loginName: string;
  principalType: "User" | "SharePointGroup" | "SecurityGroup" | "DistributionList" | "Unknown";
}

const SHAREPOINT_ROLE_IDS: Record<TilePermissionLevel, number> = {
  Read: 1073741826,
  Contribute: 1073741827,
};

const USER_PRINCIPAL_TYPES = new Set<number>([1, 4, 8]);
const TILE_GROUP_CACHE_PREFIX = "tile-access-group-memberships";

const escapeFilterValue = (value: string): string => value.replace(/'/g, "''");

const createSp = (context: WebPartContext): SPFI => spfi().using(SPFx(context));

const getMembershipCacheKey = (siteUrl: string, userId: number): string =>
  `${TILE_GROUP_CACHE_PREFIX}:${siteUrl}:${userId}`;

const readCachedGroupTitles = (cacheKey: string): string[] | null => {
  if (typeof window === "undefined") {
    return null;
  }

  try {
    const cachedValue = window.sessionStorage.getItem(cacheKey);
    if (!cachedValue) {
      return null;
    }

    const parsedValue = JSON.parse(cachedValue);
    return Array.isArray(parsedValue) ? parsedValue : null;
  } catch (error) {
    console.warn("Unable to read cached tile access memberships.", error);
    return null;
  }
};

const writeCachedGroupTitles = (cacheKey: string, groupTitles: string[]): void => {
  if (typeof window === "undefined") {
    return;
  }

  try {
    window.sessionStorage.setItem(cacheKey, JSON.stringify(groupTitles));
  } catch (error) {
    console.warn("Unable to cache tile access memberships.", error);
  }
};

const mapUserPrincipalType = (principalType?: number): ITileAccessPrincipal["principalType"] => {
  switch (principalType) {
    case 1:
      return "User";
    case 4:
      return "SecurityGroup";
    case 8:
      return "DistributionList";
    default:
      return "Unknown";
  }
};

const toUserPrincipal = (user: any): ITileAccessPrincipal => ({
  id: `${user.Id}`,
  key: `${user.Id}`,
  principalId: user.Id,
  text: user.Title,
  secondaryText: user.Email || user.LoginName,
  loginName: user.LoginName,
  principalType: mapUserPrincipalType(user.PrincipalType),
});

const toGroupPrincipal = (group: any): ITileAccessPrincipal => ({
  id: `${group.Id}`,
  key: `${group.Id}`,
  principalId: group.Id,
  text: group.Title,
  secondaryText: group.Description || group.LoginName,
  loginName: group.LoginName,
  principalType: "SharePointGroup",
});

const sortPrincipals = (items: ITileAccessPrincipal[]): ITileAccessPrincipal[] =>
  [...items].sort((left, right) => (left.text || "").localeCompare(right.text || ""));

export const getSharePointRoleDefinitionId = (permissionLevel: TilePermissionLevel): number =>
  SHAREPOINT_ROLE_IDS[permissionLevel];

export const getTileAccessGroupName = (tileName: string): string => `Tile - ${tileName?.trim()}`;

export const getCurrentUserGroupTitles = async (
  context: WebPartContext,
  useCache = true
): Promise<Set<string>> => {
  const userId = context.pageContext.legacyPageContext.userId;
  const siteUrl = context.pageContext.web.absoluteUrl;

  if (!userId) {
    return new Set<string>();
  }

  const cacheKey = getMembershipCacheKey(siteUrl, userId);

  if (useCache) {
    const cachedTitles = readCachedGroupTitles(cacheKey);
    if (cachedTitles !== null) {
      return new Set(cachedTitles);
    }
  }

  const sp = createSp(context);
  const groups = await sp.web.siteUsers
    .getById(userId)
    .groups
    .select("Id", "Title")();

  const normalizedTitles = Array.from(new Set(
    (groups || [])
      .map((group) => (group.Title || "").trim().toLowerCase())
      .filter(Boolean)
  ));

  writeCachedGroupTitles(cacheKey, normalizedTitles);

  return new Set(normalizedTitles);
};

export const userHasAccessToTile = (tileName: string, groupTitles: Set<string>): boolean =>
  groupTitles.has(getTileAccessGroupName(tileName).toLowerCase());

export const searchTileAccessPrincipals = async (
  context: WebPartContext,
  searchText: string
): Promise<ITileAccessPrincipal[]> => {
  const sp = createSp(context);
  const trimmedSearchText = searchText.trim();
  const escapedSearchText = escapeFilterValue(trimmedSearchText);

  const [users, groups] = await Promise.all([
    trimmedSearchText
      ? sp.web.siteUsers
        .select("Id", "Title", "Email", "LoginName", "PrincipalType")
        .filter(
          `PrincipalType ne 5 and (startswith(Title,'${escapedSearchText}') or startswith(Email,'${escapedSearchText}'))`
        )()
      : sp.web.siteUsers.select("Id", "Title", "Email", "LoginName", "PrincipalType").top(15)(),
    trimmedSearchText
      ? sp.web.siteGroups
        .select("Id", "Title", "Description", "LoginName")
        .filter(`startswith(Title,'${escapedSearchText}')`)()
      : sp.web.siteGroups.select("Id", "Title", "Description", "LoginName").top(10)(),
  ]);

  const mappedUsers = (users || [])
    .filter((user: any) => USER_PRINCIPAL_TYPES.has(user.PrincipalType))
    .map(toUserPrincipal);
  const mappedGroups = (groups || []).map(toGroupPrincipal);

  return sortPrincipals(
    [...mappedUsers, ...mappedGroups].filter(
      (principal, index, allPrincipals) =>
        allPrincipals.findIndex((candidate) => candidate.principalId === principal.principalId) === index
    )
  );
};

export const resolveTileAccessPrincipals = async (
  context: WebPartContext,
  principalIds: number[],
  existingPrincipals: ITileAccessPrincipal[] = []
): Promise<ITileAccessPrincipal[]> => {
  if (!principalIds?.length) {
    return [];
  }

  const sp = createSp(context);
  const resolvedPrincipals = await Promise.all(
    principalIds.map(async (principalId) => {
      const existingPrincipal = existingPrincipals.find((item) => item.principalId === principalId);
      if (existingPrincipal?.loginName) {
        return existingPrincipal;
      }

      try {
        const user = await sp.web.siteUsers.getById(principalId)();
        return toUserPrincipal(user);
      } catch (userError) {
        try {
          const group = await sp.web.siteGroups.getById(principalId)();
          return toGroupPrincipal(group);
        } catch (groupError) {
          console.warn(`Unable to resolve principal ${principalId}.`, { userError, groupError });
          return null;
        }
      }
    })
  );

  return sortPrincipals(
    resolvedPrincipals.filter((principal): principal is ITileAccessPrincipal => principal !== null)
  );
};

export const createTileAccessGroup = async (
  context: WebPartContext,
  tileName: string,
  principals: ITileAccessPrincipal[]
): Promise<{ Id: number; Title: string; LoginName: string; }> => {
  const sp = createSp(context);
  const groupName = getTileAccessGroupName(tileName);

  let groupInfo: { Id: number; Title: string; LoginName: string; };

  try {
    const existingGroup = await sp.web.siteGroups.getByName(groupName)();
    throw new Error(`A SharePoint group named "${existingGroup.Title}" already exists.`);
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "";
    const errorStatus = (error as any)?.status || (error as any)?.data?.response?.status;
    const isMissingGroupError =
      errorStatus === 404 ||
      errorMessage.includes("does not exist") ||
      errorMessage.includes("404");

    if (!isMissingGroupError) {
      throw error;
    }

    groupInfo = await sp.web.siteGroups.add({
      Title: groupName,
      Description: `Access group for the "${tileName}" tile.`,
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      OnlyAllowMembersViewMembership: true,
    });
  }

  const group = sp.web.siteGroups.getById(groupInfo.Id);
  const currentUserId = context.pageContext.legacyPageContext.userId;

  if (currentUserId) {
    await group.setUserAsOwner(currentUserId);
  }

  for (const principal of principals) {
    if (!principal.loginName) {
      continue;
    }

    try {
      await group.users.add(principal.loginName);
    } catch (error) {
      console.warn(`Unable to add ${principal.loginName} to ${groupName}.`, error);
    }
  }

  return groupInfo;
};

export const syncTileAccessGroupMembers = async (
  context: WebPartContext,
  tileName: string,
  principals: ITileAccessPrincipal[]
): Promise<{ Id: number; Title: string; LoginName: string; }> => {
  const sp = createSp(context);
  const groupName = getTileAccessGroupName(tileName);
  const desiredPrincipals = sortPrincipals(
    principals.filter((principal) => !!principal?.loginName)
  );

  let groupInfo: { Id: number; Title: string; LoginName: string; };

  try {
    groupInfo = await sp.web.siteGroups.getByName(groupName)();
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "";
    const errorStatus = (error as any)?.status || (error as any)?.data?.response?.status;
    const isMissingGroupError =
      errorStatus === 404 ||
      errorMessage.includes("does not exist") ||
      errorMessage.includes("404");

    if (!isMissingGroupError) {
      throw error;
    }

    return createTileAccessGroup(context, tileName, desiredPrincipals);
  }

  const group = sp.web.siteGroups.getById(groupInfo.Id);
  const existingMembers = await getTileAccessGroupMembers(context, tileName);
  const existingLoginNames = new Set(
    existingMembers.map((member) => member.loginName.toLowerCase())
  );
  const desiredLoginNames = new Set(
    desiredPrincipals.map((member) => member.loginName.toLowerCase())
  );

  for (const principal of desiredPrincipals) {
    if (!existingLoginNames.has(principal.loginName.toLowerCase())) {
      try {
        await group.users.add(principal.loginName);
      } catch (error) {
        console.warn(`Unable to add ${principal.loginName} to ${groupName}.`, error);
      }
    }
  }

  for (const member of existingMembers) {
    if (!desiredLoginNames.has(member.loginName.toLowerCase())) {
      try {
        await group.users.removeById(member.principalId);
      } catch (error) {
        console.warn(`Unable to remove ${member.loginName} from ${groupName}.`, error);
      }
    }
  }

  return groupInfo;
};

export const getTileAccessGroupMembers = async (
  context: WebPartContext,
  tileName: string
): Promise<ITileAccessPrincipal[]> => {
  const sp = createSp(context);
  const groupName = getTileAccessGroupName(tileName);

  try {
    const members = await sp.web.siteGroups.getByName(groupName).users
      .select("Id", "Title", "Email", "LoginName", "PrincipalType")();

    return sortPrincipals((members || []).map(toUserPrincipal));
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : "";
    const errorStatus = (error as any)?.status || (error as any)?.data?.response?.status;
    const isMissingGroupError =
      errorStatus === 404 ||
      errorMessage.includes("does not exist") ||
      errorMessage.includes("404");

    if (isMissingGroupError) {
      return [];
    }

    throw error;
  }
};
