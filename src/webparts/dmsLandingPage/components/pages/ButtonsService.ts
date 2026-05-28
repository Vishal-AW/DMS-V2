import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";

export const getAllButtonsAdmin = async (context: WebPartContext) => {
    const sp = spfi().using(SPFx(context));
    return await sp.web.lists.getByTitle("DMS_Buttons").items
        .select("*")
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
            Edit: item.EditPermission,
            Read: item.ReadPermission,
        });
    });

    return await execute();
};