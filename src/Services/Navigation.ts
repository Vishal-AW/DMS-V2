import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as FluentIcons from "@fluentui/react-icons";
export interface NavItem {
    label: string;
    href: string;
}

export interface NavSection {
    id: string;
    icon?: keyof typeof FluentIcons;
    label: string;
    href?: string;
    items?: NavItem[];
}

const getUserGroups = async (context: WebPartContext): Promise<any[]> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const userId = context.pageContext.legacyPageContext?.userId;

        if (!userId) {
            return [];
        }

        const url = `${webUrl}/_api/Web/GetUserById(${userId})?$expand=Groups`;

        const response: SPHttpClientResponse = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    Accept: "application/json;odata=nometadata",
                    "odata-version": "",
                },
            },
        );

        if (response.ok) {
            const data = await response.json();
            return data.Groups || [];
        }
    } catch (error) {
        console.error("Error getting user groups:", error);
    }
    return [];
};

const buildPermissionFilter = async (context: WebPartContext): Promise<string> => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const userId = context.pageContext.legacyPageContext?.userId;

        if (!userId) {
            return "";
        }

        const groups = await getUserGroups(context);
        const usersUrl = `${webUrl}/_api/Web/siteusers`;
        const usersResponse: SPHttpClientResponse =
            await context.spHttpClient.get(
                usersUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        Accept: "application/json;odata=nometadata",
                        "odata-version": "",
                    },
                },
            );

        let dinamicurl = `Permission/Id eq ${userId}`;

        if (usersResponse.ok) {
            const usersData = await usersResponse.json();
            const userArray: any[] = [];

            usersData.value.forEach((el: any) => {
                if (el.IsShareByEmailGuestUser === false) {
                    userArray.push(el);
                }
            });

            const externaluser = userArray;
            const NonExternalUser = externaluser.filter(
                (Title) => Title.Title === "Everyone except external users",
            );

            if (NonExternalUser.length > 0) {
                dinamicurl =
                    dinamicurl + " or Permission/Id eq " + NonExternalUser[0].Id + " ";
            }
        }

        // group permissions
        for (let i = 0; i < groups.length; i++) {
            dinamicurl = dinamicurl + " or Permission/Id eq " + groups[i].Id + " ";
        }

        return dinamicurl;
    } catch (error) {
        console.error("Error building permission filter:", error);
        return "";
    }
};


const buildNavSections = (menuData: any[]): NavSection[] => {
    const parents = menuData.filter(
        (item) => item.ParentMenuIdId === null
    );

    return parents.map((parent) => {
        const children = menuData.filter(
            (child) => child.ParentMenuIdId === parent.Id
        );

        return {
            id: parent.MenuName.toLowerCase().replace(/\s+/g, "-"),
            label: parent.MenuName,
            href: children.length === 0 ? parent.URL : undefined,
            icon: parent.IconClass,
            items:
                children.length > 0
                    ? children.map((child) => ({
                        label: child.MenuName,
                        href: child.URL
                    }))
                    : undefined
        };
    });
};

// Load menu items from SharePoint with permissions
export const loadMenuItems = async (context: WebPartContext) => {
    try {
        const webUrl = context.pageContext.web.absoluteUrl;
        const permissionFilter = await buildPermissionFilter(context);

        if (!permissionFilter) {
            console.warn("No permission filter available");
            return;
        }

        const url = `${webUrl}/_api/web/lists/getByTitle('GEN_Navigation')/items?$select=*,ParentMenuId/Id,ParentMenuId/MenuName,Permission/ID,IconClass&$expand=ParentMenuId,Permission&$orderby=OrderNo&$filter=Active eq '1' and (${permissionFilter})&$top=500`;

        const response: SPHttpClientResponse = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    Accept: "application/json;odata=nometadata",
                    "odata-version": "",
                },
            },
        );

        if (response.ok) {
            const data = await response.json();
            const menuData = data.value;

            // Store all permitted menu items
            return buildNavSections(menuData);
        }
    } catch (error) {
        console.error("Error loading menu items:", error);
        return [];
    }
};