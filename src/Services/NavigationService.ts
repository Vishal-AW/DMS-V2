/* eslint-disable */
//import { WebPartContext } from "@microsoft/sp-webpart-base";
import { GetListItem, CreateItem, UpdateItem } from "../DAL/Commonfile";


export function getdata(WebUrl: string, spHttpClient: any) {
    let filter = "";

    return getMethod(WebUrl, spHttpClient, filter);
}

export async function getAllNav(WebUrl: string, spHttpClient: any, EmailId: any) {
    const filter = 'Active eq 1';
    return getMethod(WebUrl, spHttpClient, filter);
}


export function getChildMenu(WebUrl: string, spHttpClient: any, menuId: any) {
    let filter = "ParentMenuId/Id eq '" + menuId + "'";

    return getMethod(WebUrl, spHttpClient, filter);
}
export function getChildMenunew(WebUrl: string, spHttpClient: any, menuId: any) {
    let filter = "ParentMenuId/Id eq '" + menuId + "'";

    return getMethod(WebUrl, spHttpClient, filter);
}

export function getparentdata(WebUrl: string, spHttpClient: any) {
    let filter = "ParentMenuId/Id eq '" + null + "'";

    return getMethod(WebUrl, spHttpClient, filter);
}

export function getDataByID(WebUrl: string, spHttpClient: any, ID: number) {
    let filter = "ID eq " + ID;

    return getMethod(WebUrl, spHttpClient, filter);
}


async function getMethod(WebUrl: string, spHttpClient: any, filter: any) {

    let option = {
        select: "*,isParentMenu,ID,MenuName,Permission/Id,Permission/Name,URL,OrderNo,Next_Tab,ParentMenuId/Id,ParentMenuId/MenuName,External_Url,Active,IconClass",
        expand: "ParentMenuId,Permission",
        filter: filter,
        orderby: 'OrderNo',
        top: 5000
    };

    return await GetListItem(WebUrl, spHttpClient, "GEN_Navigation", option);
}

export function SaveNavigationMaster(WebUrl: string, spHttpClient: any, savedata: any) {

    return CreateItem(WebUrl, spHttpClient, "GEN_Navigation", savedata);

}


export function UpdateNavigationMaster(WebUrl: string, spHttpClient: any, savedata: any, LID: number) {

    return UpdateItem(WebUrl, spHttpClient, "GEN_Navigation", savedata, LID);

}


