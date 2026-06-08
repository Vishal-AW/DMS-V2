import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { PermissionKind } from "@pnp/sp/security";
import { isButtonPermitted, IDmsButton } from "./buttonPermissionHelper";

// export const getAllButtonsAdmin = async (context: WebPartContext) => {
//     const sp = spfi().using(SPFx(context));
//     return await sp.web.lists.getByTitle("DMS_Buttons").items
//         .select("*")
//         .orderBy("Sequence", true)();
// };

//Set to get only active buttons
export const getAllButtonsAdmin = async (context: WebPartContext) => {
    const sp = spfi().using(SPFx(context));
    return await sp.web.lists.getByTitle("DMS_Buttons").items
        .select("*")
        .filter("Active eq 1")
        .orderBy("Sequence", true)();
};

export const updateButtonItem = async (context: WebPartContext, id: number, data: any) => {
    const sp = spfi().using(SPFx(context));
    return await sp.web.lists.getByTitle("DMS_Buttons").items.getById(id).update(data);
};

export const updateButtonsBatch = async (context: WebPartContext, items: any[]) => {
    const sp = spfi().using(SPFx(context));
    const [batchedSP, execute] = sp.web.batched();
    const list = batchedSP.lists.getByTitle("DMS_Buttons");

    items.forEach(item => {
        list.items.getById(item.ID).update({
            Title: item.Title,
            InternalName: item.InternalName,
            Active: item.Active,
            Sequence: item.Sequence,
            ButtonType: item.ButtonType,
            ButtonDisplayName: item.ButtonDisplayName,
            Icons: item.Icons,
            FullControl: item.FullControl,
            Contribute: item.Contribute,
            EditPermission: item.Edit,
            ReadPermission: item.Read,
        });
    });

    return await execute();
};

export const getPermittedButtons = async (context: WebPartContext, libraryName?: string): Promise<IDmsButton[]> => {
    const sp = spfi().using(SPFx(context));

    // 1. Fetch all active buttons from DMS_Buttons list
    const allButtons = await sp.web.lists
        .getByTitle('DMS_Buttons')
        .items
        .filter('Active eq 1')
        .select('Title', 'InternalName', 'ButtonType', 'ButtonDisplayName', 'Icons', 'Sequence', 'FullControl', 'Contribute', 'EditPermission', 'ReadPermission','IsSetReadInactive')
        .orderBy('Sequence')();

    // 2. Get current user effective permissions
    const userPerms = libraryName
        ? await sp.web.lists.getByTitle(libraryName).getCurrentUserEffectivePermissions()
        : await sp.web.getCurrentUserEffectivePermissions();

    // 3. Resolve user permission level
    const hasFullControl = sp.web.hasPermissions(userPerms, PermissionKind.ManagePermissions);
    const hasContribute = sp.web.hasPermissions(userPerms, PermissionKind.AddListItems);
    const hasEdit = sp.web.hasPermissions(userPerms, PermissionKind.EditListItems);
    const hasRead = sp.web.hasPermissions(userPerms, PermissionKind.ViewListItems);

    // 4. Filter buttons based on permission level
    return allButtons.filter(btn =>
        isButtonPermitted(btn, {
            hasFullControl, hasContribute, hasEdit, hasRead
        })
    );
};

