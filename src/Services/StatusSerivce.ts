import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';

export function GetAllStatus(WebUrl: string, spHttpClient: any) {
  let filter = "";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getStatusByInternalStatus(WebUrl: string, spHttpClient: any, internalStatus: string) {
  let filter = "InternalStatus eq '" + internalStatus + "'";

  return getMethod(WebUrl, spHttpClient, filter);
}


async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,StatusName,InternalStatus",
    //expand:"Designation,Status,Manager,HOD",
    filter: filter,
    orderby: 'StatusName',
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "DMS_Mas_Status", option);
}


export function SaveStateMaster(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "DMS_Mas_Status", savedata);

}


export function UpdateStateMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "DMS_Mas_Status", savedata, LID);

}