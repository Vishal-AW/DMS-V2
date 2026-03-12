import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";


// export const FolderStructure = async (context: WebPartContext, FolderPath: string, uid: number[], LibraryName: string) => {

//     const folderUrl = `${context.pageContext.web.absoluteUrl}/${FolderPath}`;
//     return await context.spHttpClient.post(
//         `${context.pageContext.web.absoluteUrl}/_api/web/folders?$expand=ListItemAllFields`,
//         SPHttpClient.configurations.v1,
//         {
//             headers: {
//                 Accept: "application/json;odata=nometadata",
//                 "odata-version": "3.0",
//                 "X-HTTP-Method": "POST",
//                 "Content-Type": "application/json",
//             },
//             body: JSON.stringify({ ServerRelativeUrl: folderUrl }),
//         }
//     ).then(async (response: SPHttpClientResponse) => {
//         if (response.ok) {
//             const data = await response.json();
//             await breakRoleInheritance(context, FolderPath, uid, LibraryName, data.ListItemAllFields.ID);
//             return data.ListItemAllFields.ID;
//         }
//     }).catch((error) => {
//         console.error('Error creating folder:', error);
//     });
// };

export const FolderStructure = async (context: WebPartContext, FolderPath: string, uid: number[], LibraryName: string, ChildFolderRoleInheritance: boolean) => {

    const folderUrl = `${context.pageContext.web.absoluteUrl}/${FolderPath}`;
    return await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/web/folders?$expand=ListItemAllFields`,
        SPHttpClient.configurations.v1,
        {
            headers: {
                Accept: "application/json;odata=nometadata",
                "odata-version": "3.0",
                "X-HTTP-Method": "POST",
                "Content-Type": "application/json",
            },
            body: JSON.stringify({ ServerRelativeUrl: folderUrl }),
        }
    ).then(async (response: SPHttpClientResponse) => {
        if (response.ok) {
            const data = await response.json();
            if (ChildFolderRoleInheritance) {
                await breakRoleInheritance(context, FolderPath, uid, LibraryName, data.ListItemAllFields.ID);
            }
            return data.ListItemAllFields.ID;
        }
    }).catch((error) => {
        console.error('Error creating folder:', error);
    });
};
const breakRoleInheritance = async (context: WebPartContext, folderUrl: string, userIds: number[], LibraryName: string, Id: number) => {

    const breakInheritanceUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/breakroleinheritance(true)`;
    return await context.spHttpClient.post(
        breakInheritanceUrl,
        SPHttpClient.configurations.v1,
        {
            headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
            },
        }
    ).then(async (response: SPHttpClientResponse) => {
        if (response.ok) {
            await grantPermissions(context, folderUrl, [...userIds]);
            return await removeAllPermissions(context, folderUrl, [...userIds]);



        }
    });

};

// const grantPermissions = async (context: WebPartContext, folderUrl: string, userIds: number[]) => {
//     try {
//         for (const userId of userIds) {
//             const permissionUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${userId},roleDefId=1073741827)`;
//             const response = await context.spHttpClient.post(
//                 permissionUrl,
//                 SPHttpClient.configurations.v1,
//                 {
//                     headers: {
//                         Accept: 'application/json;odata=verbose',
//                         'Content-Type': 'application/json;odata=verbose',
//                     },
//                 }
//             );

//             if (!response.ok) {
//                 console.error('Failed to grant permission for user ID:', userId);
//             }
//         }
//     } catch (error) {
//         console.error(error);
//     }
// };

const grantPermissions = async (context: WebPartContext, folderUrl: string, userIds: any[]) => {
    try {
        for (const userId of userIds) {
            if (userId.type === "FolderAccess") {
                const permissionUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${userId.id},roleDefId=1073741827)`;
                const response = await context.spHttpClient.post(
                    permissionUrl,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            Accept: 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose',
                        },
                    }
                );

                if (!response.ok) {
                    console.error('Failed to grant permission for user ID:', userId);
                }
            }
            else if (userId.type === "Admin" || userId.type === "TileAdmin") {
                const permissionUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${userId.id},roleDefId=1073741829)`;
                await context.spHttpClient.post(
                    permissionUrl,
                    SPHttpClient.configurations.v1,
                    {
                        headers: {
                            Accept: 'application/json;odata=verbose',
                            'Content-Type': 'application/json;odata=verbose',
                        },
                    }
                );

                // if (!response.ok) {
                //     console.error('Failed to grant permission for user ID:', userId);
                // }
                // else {

                //     const permissionUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments/addroleassignment(principalid=${userId.id},roleDefId=1073741827)`;
                //     const response = await context.spHttpClient.post(
                //         permissionUrl,
                //         SPHttpClient.configurations.v1,
                //         {
                //             headers: {
                //                 Accept: 'application/json;odata=verbose',
                //                 'Content-Type': 'application/json;odata=verbose',
                //             },
                //         }
                //     );

                //     if (!response.ok) {
                //         console.error('Failed to grant permission for user ID:', userId);
                //     }
                // }
            }
        }
    } catch (error) {
        console.error(error);
    }
}; const removeAllPermissions = async (context: WebPartContext, folderUrl: string, userIds: any[]) => {
    const roleAssignmentsUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderUrl}')/ListItemAllFields/roleassignments`;

    try {
        const response = await context.spHttpClient.get(
            roleAssignmentsUrl,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }
        );

        if (!response.ok) {
            console.error('Failed to fetch role assignments:', response.statusText);
            return;
        }

        const data = await response.json();
        const valuedata = data.value;

        // Filter only the roles where PrincipalId is in the userIds array
        //const rolesToRemove = roleAssignments.filter((role: any) => userIds.id.includes(role.PrincipalId));
        const userIds1 = (userIds as any[]).map(role => role.id);
        const filtered = valuedata.filter(
            (role: any) => !userIds1.includes(role.PrincipalId)
        );

        console.log(filtered);
        for (const role of filtered) {
            const deleteUrl = `${roleAssignmentsUrl}/removeroleassignment(principalid=${role.PrincipalId})`;

            const deleteResponse = await context.spHttpClient.post(
                deleteUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        'Accept': 'application/json;odata=nometadata',
                        'Content-Type': 'application/json;odata=nometadata',
                    },
                }
            );

            if (!deleteResponse.ok) {
                console.error(`Failed to remove role assignment for PrincipalId ${role.PrincipalId}`);
            } else {
                console.log(`Successfully removed role assignment for PrincipalId ${role.PrincipalId}`);
            }
        }
    } catch (error) {
        console.error('Error in removeAllPermissions:', error);
    }
};

export const breakRoleInheritanceForLib = async (context: WebPartContext, libName: string, userIds: any[]) => {

    const breakInheritanceUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libName}')/breakroleinheritance(true)`;
    return await context.spHttpClient.post(
        breakInheritanceUrl,
        SPHttpClient.configurations.v1,
        {
            headers: {
                Accept: 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
            },
        }
    ).then(async (response: SPHttpClientResponse) => {
        if (response.ok) {
            await grantPermissionsForLib(context, libName, [...userIds]);
            return await removeAllPermissionsForLib(context, libName, [...userIds]);
        }
    });

};


// const removeAllPermissionsForLib = async (context: WebPartContext, libName: string, userIds: number[]) => {
//     const roleAssignmentsUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libName}')/roleassignments`;

//     try {
//         return await context.spHttpClient.get(roleAssignmentsUrl,
//             SPHttpClient.configurations.v1,
//             {
//                 headers: {
//                     'Accept': 'application/json;odata=nometadata',
//                     'odata-version': ''
//                 }
//             }).then(async (response: SPHttpClientResponse) => {
//                 if (response.ok) {
//                     const data = await response.json();
//                     for (const assignment of data.value) {
//                         if (!userIds.includes(assignment.PrincipalId)) {
//                            const deleteUrl = `${roleAssignmentsUrl}/removeroleassignment(principalid=${assignment.PrincipalId})`;

//                             const deleteResponse = await context.spHttpClient.post(
//                                 deleteUrl,
//                                 SPHttpClient.configurations.v1,
//                                 {
//                                     headers: {
//                                         Accept: "application/json;odata=nometadata", // Consistent header value
//                                         "Content-Type": "application/json;odata=nometadata",
//                                     },
//                                 }
//                             );

//                             if (!deleteResponse.ok) {
//                                 console.error('Failed to remove role assignment:', assignment.PrincipalId);
//                             }
//                         }
//                     }
//                 } else {
//                     console.error('Failed to fetch role assignments:', response.statusText);
//                 }
//             }).catch((err: any) => {
//                 console.log(err);
//             });


//     } catch (error) {
//         console.error('Error in removeAllPermissions:', error);
//     }
// };
const removeAllPermissionsForLib = async (context: WebPartContext, libName: string, userIds: number[]) => {
    const roleAssignmentsUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${libName}')/roleassignments`;

    try {
        return await context.spHttpClient.get(roleAssignmentsUrl,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'odata-version': ''
                }
            }).then(async (response: SPHttpClientResponse) => {
                if (response.ok) {
                    const data = await response.json();
                    const valuedata = data.value;

                    const userIds1 = (userIds as any[]).map(role => role.IDs);


                    const filtered = valuedata.filter(
                        (role: any) => !userIds1.includes(role.PrincipalId)
                    );

                    console.log(filtered);


                    for (const role of filtered) {
                        const deleteUrl = `${roleAssignmentsUrl}/removeroleassignment(principalid=${role.PrincipalId})`;

                        const deleteResponse = await context.spHttpClient.post(
                            deleteUrl,
                            SPHttpClient.configurations.v1,
                            {
                                headers: {
                                    Accept: "application/json;odata=nometadata",
                                    "Content-Type": "application/json;odata=nometadata",
                                },
                            }
                        );

                        if (!deleteResponse.ok) {
                            console.error('Failed to remove role assignment:', role.PrincipalId);
                        }
                    }


                } else {
                    console.error('Failed to fetch role assignments:', response.statusText);
                }
            }).catch((err: any) => {
                console.log(err);
            });


    } catch (error) {
        console.error('Error in removeAllPermissions:', error);
    }
};

export const grantPermissionsForLib = async (context: WebPartContext, libName: string, userIds: any[]) => {
    try {
        for (const userId of userIds) {
            const permissionType = userId.Type === "User" ? 1073741827 : 1073741829;

            const permissionUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${libName}')/roleassignments/addroleassignment(principalid=${userId.IDs},roleDefId=${permissionType})`;
            const response = await context.spHttpClient.post(
                permissionUrl,
                SPHttpClient.configurations.v1,
                {
                    headers: {
                        Accept: 'application/json;odata=verbose',
                        'Content-Type': 'application/json;odata=verbose',
                    },
                }
            );

            if (!response.ok) {
                console.error('Failed to grant permission for user ID:', userId);
            }
        }
    } catch (error) {
        console.error(error);
    }
};