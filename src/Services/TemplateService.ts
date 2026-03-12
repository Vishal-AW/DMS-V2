import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';

export function getTemplate(WebUrl: string, spHttpClient: any) {
  let filter = "";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getTemplateActive(WebUrl: string, spHttpClient: any) {
  let filter = "Active eq 1";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getTemplateDataByID(WebUrl: string, spHttpClient: any, ID: number) {
  let filter = "ID eq " + ID;

  return getMethod(WebUrl, spHttpClient, filter);
}



async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,Name,Active",
    filter: filter,
    orderby: 'Name',
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "DMS_Template", option);
}


export function SaveTemplateMaster(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "DMS_Template", savedata);

}


export function UpdateTemplateMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "DMS_Template", savedata, LID);

}