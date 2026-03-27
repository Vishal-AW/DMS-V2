
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { GetListItem, UpdateItem } from "../DAL/Commonfile";

import { SPPermission } from '@microsoft/sp-page-context';

const PermissionMap: Record<string, any> = {
    viewListItems: SPPermission.viewListItems,
    openItems: SPPermission.openItems,
    viewVersions: SPPermission.viewVersions,
    viewFormPages: SPPermission.viewFormPages,
    editListItems: SPPermission.editListItems,
    deleteListItems: SPPermission.deleteListItems,
    approveItems: SPPermission.approveItems,
    // add more if needed…
};

export async function getAllFolder(WebUrl: string, context: WebPartContext, FolderName: string) {
    const url = WebUrl + "/_api/Web/GetFolderByServerRelativeUrl('" + FolderName + "')?$select=*&$orderby=Id desc&$expand=Files/CheckedOutByUser,Folders,Files,Files/ModifiedBy,Folders/ListItemAllFields,Files/ListItemAllFields,ListItemAllFields,Files/Status,FileLeafRef,FileRef,FileDirRef";


    return await context.spHttpClient.get(url,
        SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
            }
        }).then(async (response: SPHttpClientResponse) => {
            return response.json();
        }).catch((err: any) => {
            console.log(err);
        });

}



export async function getPermission(url: string, context: WebPartContext) {

    return await context.spHttpClient.get(url,
        SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
            }
        }).then(async (response: SPHttpClientResponse) => {
            return response.json();
        }).catch((err: any) => {
            console.log(err);
        });

}

export async function commonPostMethod(url: string, context: WebPartContext) {
    return await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': '3.0',
            'X-HTTP-Method': 'POST'
        }
    }).then((response: SPHttpClientResponse) => {
        if (response.ok) {
            return response;
        }
    });
}

export async function getListData(url: string, context: WebPartContext) {

    return await context.spHttpClient.get(url,
        SPHttpClient.configurations.v1,
        {
            headers: {
                'Accept': 'application/json;odata=nometadata',
                'odata-version': ''
            }
        }).then(async (response: SPHttpClientResponse) => {
            return response.json();
        }).catch((err: any) => {
            console.log(err);
        });
}

//Created by rupali
export async function checkUserIsSiteAdminById(context: any, userId: number) {
    try {
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/siteusers/getbyid(${userId})?$select=Id,LoginName,IsSiteAdmin`;

        const response = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) {
            console.error("Error fetching user");
            return false;
        }

        const user = await response.json();
        return user.IsSiteAdmin; // true / false
    } catch (error) {
        console.error("Error:", error);
        return false;
    }
}

//Created by Rupali
export async function checkUserInProjectAdmin(context: any, userId: number) {
    try {
        const siteUrl = context.pageContext.web.absoluteUrl;

        const url = `${siteUrl}/_api/web/sitegroups/getbyname('ProjectAdmin')/users?$select=Id,Title`;

        const response = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1
        );

        if (!response.ok) {
            console.error("Error fetching ProjectAdmin group users");
            return false;
        }

        const data = await response.json();

        // Check if userId exists in group
        const isUserPresent = data.value.some((user: any) => user.Id === userId);

        return isUserPresent; // true / false
    } catch (error) {
        console.error("Error:", error);
        return false;
    }
}

//new added by rupali to check is user restricted view permision

const normalizeFolderServerRelativeUrl = (context: WebPartContext, folderPath: string): string => {
    if (!folderPath) return "";

    if (folderPath.startsWith("/")) {
        return folderPath;
    }

    const webRelativeUrl = context.pageContext.web.serverRelativeUrl;
    return webRelativeUrl === "/"
        ? `/${folderPath}`
        : `${webRelativeUrl}/${folderPath}`;
};

export const hasFolderPermission = async (
    context: any,
    folderServerRelativeUrl: string,
    permissionKind: keyof SPPermission
): Promise<boolean> => {
    try {
        const normalizedFolderPath = normalizeFolderServerRelativeUrl(context, folderServerRelativeUrl);
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/getFolderByServerRelativeUrl('${normalizedFolderPath}')/ListItemAllFields?$select=EffectiveBasePermissions`;

        const response = await context.spHttpClient.get(url,
            SPHttpClient.configurations.v1);

        const json = await response.json();

        if (!json.EffectiveBasePermissions) {
            return false;
        }

        const perm = new SPPermission({
            High: json.EffectiveBasePermissions.High,
            Low: json.EffectiveBasePermissions.Low
        });

        return perm.hasPermission(PermissionMap[permissionKind]);

    } catch (error) {
        console.error("Permission check error:", error);
        return false;
    }
};


export function updateLibrary(WebUrl: string, spHttpClient: SPHttpClient, metaData: any, Id: number, listName: string) {
    return UpdateItem(WebUrl, spHttpClient, listName, metaData, Id);
}

export async function UploadFile(WebUrl: string, spHttpClient: any, file: string, DisplayName: string | File, DocumentLib: string, jsonBody: any, FolderPath: string): Promise<any> {
    // let fileupload = FolderPath +"/"+FolderName;
    return new Promise((resolve) => {
        const spOpts: ISPHttpClientOptions = {
            body: file
        };
        var redirectionURL = WebUrl + "/_api/Web/GetFolderByServerRelativeUrl('" + FolderPath + "')/Files/Add(url='" + DisplayName + "', overwrite=true)?$expand=ListItemAllFields";
        spHttpClient.post(redirectionURL, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
            response.json().then(async (responseJSON: any) => {
                // console.log(responseJSON.ListItemAllFields.ID);
                // var serverRelURL = await responseJSON.ServerRelativeUrl;
                if (jsonBody != null) {
                    let IsExistingRefID = responseJSON.ListItemAllFields.ID.toString();
                    if (jsonBody.IsExistingRefID !== null && jsonBody.IsExistingRefID !== undefined && jsonBody.IsExistingRefID !== "") {
                        IsExistingRefID = jsonBody.IsExistingRefID;
                    }
                    jsonBody.IsExistingRefID = IsExistingRefID;
                    await UpdateItem(WebUrl, spHttpClient, DocumentLib, jsonBody, responseJSON.ListItemAllFields.ID);

                }
                resolve(responseJSON);
            });
        });
    });

}

export function getApprovalData(context: WebPartContext, libeName: string, useremail: string) {
    const filter = "CurrentApprover eq '" + useremail + "' and Active eq 1";
    return getMethod(context.pageContext.web.absoluteUrl, context.spHttpClient, filter, libeName);
}

export function getRecycleData(context: WebPartContext, libeName: string) {
    const filter = "Active eq 0";
    return getMethod(context.pageContext.web.absoluteUrl, context.spHttpClient, filter, libeName);
}

export function getArchiveData(context: WebPartContext, libeName: string) {
    const filter = "Active eq 0 and IsArchiveFlag eq 1 and DeleteFlag ne 'Deleted'";
    return getDocument(context.pageContext.web.absoluteUrl, context.spHttpClient, filter, libeName);
}

export function getDataByRefID(context: WebPartContext, Id: string, libeName: string) {
    const filter = "IsExistingRefID eq " + Id;
    return getDocument(context.pageContext.web.absoluteUrl, context.spHttpClient, filter, libeName);
}

async function getMethod(WebUrl: string, spHttpClient: any, filter: any, libeName: string) {

    let option = {
        select: "*,Projectmanager/Id,Projectmanager/Title,Publisher/Id,Publisher/Title,Status/Id,Status/StatusName,Author/EMail,Author/Title",
        expand: "File,Projectmanager,Publisher,Status,Author",
        filter: filter,
        orderby: 'ID desc',
        top: 5000
    };

    return await GetListItem(WebUrl, spHttpClient, libeName, option);
}


export async function getDocument(WebUrl: string, spHttpClient: any, filter: any, libName: string) {

    var selectcols = "*,ID,File,DefineRole,ProjectmanagerAllow,Projectmanager/Id,Projectmanager/Title,ProjectmanagerEmail,PublisherAllow,Publisher/Id,";
    selectcols += "Publisher/Title,PublisherEmail,CurrentApprover,InternalStatus,ProjectMasterLID,";
    selectcols += "LatestRemark,AllowApprover,Created,Author/EMail,Author/Title,FileLeafRef,FileRef,FileDirRef,Active,ProjectmanagerId,PublisherId,File,ServerRedirectedEmbedUrl,DisplayStatus,Level,OCRStatus,";
    selectcols += "Company,Template,IsArchiveFlag";
    var option = {
        select: selectcols,
        expand: "File,Projectmanager,Publisher,Author",
        filter: filter,
        orderby: 'ID desc',
        top: 5000
    };


    return await GetListItem(WebUrl, spHttpClient, libName, option);
}


export function generateAutoRefNumber(refCount: any, data: any, CreatedDate: any, libDetails: any) {
    let refNo = "";
    let incrementCount = 0;
    const currentFY: any = getFinancialYear(new Date());
    if (libDetails.IsDynamicReference) {
        const refFormulaValue = libDetails.ReferenceFormula.split(libDetails.Separator);
        refFormulaValue.map(function (el: any, i: number) {
            const pattern = /\{(.*?)\}/g;
            const matches = el.match(pattern);
            if (matches == null)
                refNo += `${el}${libDetails.Separator}`;
            else {
                matches.map(function (element: any, ind: number) {
                    const elementId = element.replace(/[^a-z0-9\s-_]/gi, '');
                    if (refFormulaValue.length - 1 == i && matches.length - 1 == ind) {
                        incrementCount = initialIncrement(elementId, refCount, CreatedDate);
                        refNo += padLeft(incrementCount.toString(), 5, "0");
                    } else {
                        if (elementId == "YY_YY")
                            refNo += `${currentFY.startYear}${libDetails.Separator}${currentFY.endYear}`;
                        else if (elementId == "YYYY")
                            refNo += `${new Date().getFullYear()}`;
                        else if (elementId == "MM")
                            refNo += `${new Date().toLocaleString('default', { month: '2-digit' })}`;
                        else
                            refNo += `${data[elementId]}`;
                    }
                });
                refFormulaValue.length - 1 != i ? (refNo += libDetails.Separator) : "";
            }
        });
    } else {
        incrementCount = refCount > 0 ? (refCount + 1) : 1;
        const year = new Date().getFullYear();
        refNo = year + '-' + padLeft(incrementCount.toString(), 5, "0");
    }
    const obj = { "refNo": refNo, "count": incrementCount };
    return obj;
}

function initialIncrement(val: any, incrementCount: any, CreatedDate: any) {
    const lastMonth = new Date(CreatedDate).toLocaleString('default', { month: '2-digit' });
    const lastYear = new Date(CreatedDate).getFullYear();
    const month = new Date().toLocaleString('default', { month: '2-digit' });
    const year = new Date().getFullYear();
    const FY: any = getFinancialYear(new Date());
    const lastFY: any = getFinancialYear(new Date(CreatedDate));
    switch (val) {
        case "Continue":
            return incrementCount > 0 ? (incrementCount + 1) : 1;
        case "Monthly":
            return lastMonth == month ? (incrementCount + 1) : 1;
        case "Yearly":
            return lastYear == year ? (incrementCount + 1) : 1;
        case "FinancialYear":
            lastFY.endYear == FY.endYear ? (incrementCount + 1) : 1;
            break;
    }
}

function getFinancialYear(date: any) {
    const today = date;
    const fn: any = {};
    const year = today.toLocaleString('default', { year: '2-digit' });
    if ((today.getMonth() + 1) <= 3) {
        fn.startYear = (Number(year) - 1).toString();
        fn.endYear = year;
    } else {
        fn.startYear = year;
        fn.endYear = (Number(year) + 1).toString();
    }
    return fn;
}

function padLeft(value: string, length: number, char: string = "0"): string {
    return char.repeat(Math.max(0, length - value.length)) + value;
}



export async function checkPermissions(context: any, folderPath: string): Promise<boolean> {
    try {
        // const url = `${context.pageContext.web.absoluteUrl}/_api/web/DoesUserHavePermissions?high=${permissionMaskHigh}&low=${permissionMaskLow}`;

        const normalizedFolderPath = normalizeFolderServerRelativeUrl(context, folderPath);
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${normalizedFolderPath}')/ListItemAllFields/effectiveBasePermissions`;

        const response: SPHttpClientResponse = await context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }
        );

        if (!response.ok) {
            throw new Error(`Failed to get folder permissions: ${response.statusText}`);
        }
        const data = await response.json();
        const high = parseInt(data.High, 10);
        const low = parseInt(data.Low, 10);

        const readListItems = 1; // Read permission bit
        const addListItems = 3;  // Add permission bit
        const editListItems = 6; // Edit permission bit
        const deleteListItems = 7; // Delete permission bit

        const hasPermission = (bit: number): boolean => {
            if (bit < 32) {
                return (low & (1 << bit)) !== 0;
            } else {
                return (high & (1 << (bit - 32))) !== 0;
            }
        };

        const canRead = hasPermission(readListItems);
        const canAdd = hasPermission(addListItems);
        const canEdit = hasPermission(editListItems);
        const canDelete = hasPermission(deleteListItems);
        console.log("User canRead:", canRead);
        console.log("User canAdd:", canAdd);
        console.log("User canEdit:", canEdit);
        console.log("User canDelete:", canDelete);

        // Define hasWriteAccess based on Add or Edit permissions
        const hasWriteAccess = canAdd || canEdit;
        console.log("User has write access:", hasWriteAccess);
        let checkData = false;

        // If user has Read but no write and no delete access
        if (canRead === true) {
            if (canAdd === true || canEdit === true) {
                checkData = true; // Delete-only (if you want to treat this as read-only)
            } else if (canDelete) {
                checkData = false; // Has write access
            }
        } else {
            checkData = false; // Read-only
        }
        return checkData;

    } catch (error) {
        console.error("Error checking permissions:", error);
        return false;
    }
}


// export async function checkPermissions(context: any, folderPath: string): Promise<boolean> {
//     try {
//         const url = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/ListItemAllFields/RoleAssignments?$expand=RoleDefinitionBindings`;

//         const response: SPHttpClientResponse = await context.spHttpClient.get(
//             url,
//             SPHttpClient.configurations.v1,
//             {
//                 headers: {
//                     Accept: "application/json;odata=nometadata"
//                 }
//             }
//         );

//         if (!response.ok) {
//             throw new Error(`Error fetching permissions: ${response.statusText}`);
//         }

//         const data = await response.json();
//         const roleAssignments = data.value || [];
//         const hasRequiredPermissions = roleAssignments.some((assignment: any) => {
//             return assignment.RoleDefinitionBindings.some((role: any) =>
//                 ["Edit", "Contribute", "Full Control"].includes(role.Name)
//             );
//         });

//         return hasRequiredPermissions;
//     } catch (error) {
//         console.error("Error checking permissions:", error);
//         return false;
//     }
// };
