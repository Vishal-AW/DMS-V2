import { GetListItem } from '../DAL/Commonfile';

export function getRoles(WebUrl: string, spHttpClient: any) {
  let filter = "Active eq 1";

  return getMethod(WebUrl, spHttpClient, filter);
}



async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

  let option = {
    select: "ID,Title",
    //expand:"Designation,Status,Manager,HOD",
    filter: filter,
    orderby: "Title asc",
    top: 5000
  };

  return await GetListItem(WebUrl, spHttpClient, "DMS_RoleMaster", option);
}



