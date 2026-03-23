import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';

export function getArchiveDaysDetails(WebUrl: string, spHttpClient: any) {
  let filter = "";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getRedundancyDaysByID(WebUrl: string, spHttpClient: any, ID: number) {
  let filter = "ID eq " + ID;

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getActiveRedundancyDays(WebUrl: string, spHttpClient: any) {
  var filter = "Active eq '1'";

  return getMethod(WebUrl, spHttpClient, filter);
}

async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,RedundancyDays,Active",
    //expand:"Designation,Status,Manager,HOD",
    filter: filter,
    orderby: 'RedundancyDays',
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "ArchiveRedundancyDay", option);
}


export function SaveClassificationMaster(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "ArchiveRedundancyDay", savedata);

}


export function UpdateClassificationMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "ArchiveRedundancyDay", savedata, LID);

}