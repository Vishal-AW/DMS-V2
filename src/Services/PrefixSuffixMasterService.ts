import { GetListItem } from "../DAL/Commonfile";

export function getActiveTypeData(WebUrl: string, spHttpClient: any, PSType: string) {
    var filter = `Active eq '1' and PSType eq '${PSType}'`;
    return getMethod(WebUrl, spHttpClient, filter);
}
async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

    let option = {
        select: "ID,PSName,PSType,Active",
        filter: filter,
        top: 5000,
        orderby: "Id desc"
    };

    return await GetListItem(WebUrl, spHttpClient, "PrefixSuffixMaster", option);
}
