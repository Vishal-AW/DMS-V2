
import { GetListItem } from '../DAL/Commonfile';

export function getAllButtons(WebUrl: string, spHttpClient: any) {
  let filter = "Active eq 1";

  return getMethod(WebUrl, spHttpClient, filter);
}



async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,Title,ButtonType,Sequence,Active,InternalName,ButtonDisplayName,Icons",
    orderby: "Sequence asc",
    filter: filter,
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "DMS_Buttons", option);
}



