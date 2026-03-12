import { CreateItem, GetListItem } from "../DAL/Commonfile";
export function getHistoryByID(WebUrl: string, spHttpClient: any, ID: number, libName: string) {
    let filter = "DocumetLID eq " + ID + " and LibName eq '" + libName + "'";

    return getMethod(WebUrl, spHttpClient, filter);
}

async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

    let option = {
        select: "Id,DocumetLID,ActionDate,Author/Title,Action,InternalComment",
        expand: "Author",
        filter: filter,
        orderby: "Id desc",
        top: 5000
    };

    return await GetListItem(WebUrl, spHttpClient, "DMS_GeneralDocumentHistory", option);
}
export function createHistoryItem(WebUrl: string, spHttpClient: any, savedata: any) {
    return CreateItem(WebUrl, spHttpClient, "DMS_GeneralDocumentHistory", savedata);
}