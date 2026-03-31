import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';

export function getParent(WebUrl: string, spHttpClient: any) {
  const filter = "";
  return getMethod(WebUrl, spHttpClient, filter);
}
export function getActiveFolder(WebUrl: string, spHttpClient: any) {
  const filter = "Active eq 1";
  return getMethod(WebUrl, spHttpClient, filter);
}

export function getfolders(WebUrl: string, spHttpClient: any) {
  const filter = "IsParentFolder eq 1 and Active eq 1";
  return getMethod(WebUrl, spHttpClient, filter);
}

export function getFolderDataByID(WebUrl: string, spHttpClient: any, ID: number) {
  const filter = "ParentFolderIdId eq " + ID;
  return getMethod(WebUrl, spHttpClient, filter);
}

export function getChildDataByID(WebUrl: string, spHttpClient: any, ID: number) {
  const filter = "ID eq " + ID;
  return getMethod(WebUrl, spHttpClient, filter);
}

export function getTemplateDataByID(WebUrl: string, spHttpClient: any, TemplateId: number) {
  const filter = "TemplateName/ID eq " + TemplateId;
  return getMethod(WebUrl, spHttpClient, filter);
}



async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {
  const option = {
    select: "*,ID,FolderName,ParentFolderIdId,ParentFolderId/Id,ParentFolderId/FolderName,TemplateName/Name,TemplateName/ID,Active,IsParentFolder,IsApproverFlow",
    expand: "ParentFolderId,TemplateName",
    filter: filter,
    top: 5000,
    orderby: "ID desc"
  };
  return await GetListItem(WebUrl, spHttpClient, "DMS_Mas_FolderMaster", option);
}


export function SaveFolderMaster(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "DMS_Mas_FolderMaster", savedata);

}


export function UpdateFolderMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "DMS_Mas_FolderMaster", savedata, LID);

}

