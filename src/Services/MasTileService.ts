import { GetListItem, CreateItem, UpdateItem } from '../DAL/Commonfile';
import { SPHttpClient } from '@microsoft/sp-http';

export function getTileAllData(WebUrl: string, spHttpClient: SPHttpClient) {
  let filter = "";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getDataById(WebUrl: string, spHttpClient: any, ID: number) {
  let filter = "ID eq " + ID;

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getAllActiveTileData(WebUrl: string, spHttpClient: any) {
  let filter = "Active eq 1";

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getTileAdmin(WebUrl: string, spHttpClient: any, ID: number) {
  let filter = "TileAdmin/Id eq " + ID;

  return getMethod(WebUrl, spHttpClient, filter);
}
export function getDataByLibraryName(WebUrl: string, spHttpClient: any, name: any) {
  let filter = `LibraryName eq '${name}'`;

  return getMethod(WebUrl, spHttpClient, filter);
}


async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,TileName,TileImageURL,SystemCreated,Permission/Name,Permission/Id,Permission/Title,Documentpath,Active,Order0,AllowApprover,Editor/Title,Modified,LibraryName,LibGuidName,AllowOrder,DynamicControl,IsDynamicReference,ReferenceFormula,Separator,ShowMoreActions,IsArchiveRequired,RetentionDays,ArchiveVersionCount,ArchiveLibraryName,TileAdmin/Id,TileAdmin/Title,TileAdmin/EMail,IsAllowFieldsInFile,CustomPermission",
    expand: "Permission,Editor,TileAdmin",
    filter: filter,
    orderby: "ID desc",
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "DMS_Mas_Tile", option);
}


export function SaveTileSetting(WebUrl: string, spHttpClient: any, savedata: any) {

  return CreateItem(WebUrl, spHttpClient, "DMS_Mas_Tile", savedata);

}


export function UpdateTileSetting(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

  return UpdateItem(WebUrl, spHttpClient, "DMS_Mas_Tile", savedata, LID);

}