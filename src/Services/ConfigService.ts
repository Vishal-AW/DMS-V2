import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';

export function getConfig(WebUrl: string, spHttpClient: any) {
  let filter = "";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getConfigActive(WebUrl: string, spHttpClient: any) {
  let filter = "IsActive eq 1";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getConfidDataByID(WebUrl: string, spHttpClient: any, ID: number) {
  let filter = "ID eq " + ID;

  return getMethod(WebUrl, spHttpClient, filter);
}



async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "Id,Title,ColumnType,InternalListName,IsActive,IsStaticValue,StaticDataObject,DisplayValue,InternalTitleName,IsShowAsFilter,Abbreviation",
    //expand : "",
    filter: filter,
    orderby: "Id desc",
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "ConfigEntryMaster", option);
}


export function SaveconfigMaster(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "ConfigEntryMaster", savedata);

}


export function UpdateconfigMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "ConfigEntryMaster", savedata, LID);

}