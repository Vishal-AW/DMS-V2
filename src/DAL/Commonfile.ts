/* eslint-disable */
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';



export function getUSERID(WebUrl: string, spHttpClient: SPHttpClient, username: any) {

  let url = WebUrl + "/_api/web/SiteUserInfoList/items?$select=Id&$filter=Title eq '" + username + "'";


  return spHttpClient.get(url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    }).then((response: SPHttpClientResponse) => {
      console.log("response");
      /*if(response.ok)
      {
         response.json().then((data)=>{
             console.log(data.value);
             var userID=data;
        alert(userID);
        });
      }*/

      return response.json();
    });
}

export function GetListItem(WebUrl: string, spHttpClient: SPHttpClient, ListName: string, options: any) {

  //let returnval =[];
  let url = WebUrl + "/_api/web/lists/getbytitle('" + ListName + "')/Items";
  url = URLBuilder(url, options);

  return spHttpClient.get(url,
    SPHttpClient.configurations.v1,
    {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    }).then((response: SPHttpClientResponse) => {
      console.log("response");

      return response.json();
    });

}


export function CreateItem(WebUrl: string, spHttpClient: SPHttpClient, ListName: string, jsonBody: any) {



  if (!jsonBody.__metadata) {
    jsonBody.__metadata = {
      'type': 'SP.ListItem'
    };
  }


  const URL = WebUrl + "/_api/web/lists/getbytitle('" + ListName + "')/Items";
  return spHttpClient.post(URL,
    SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0'
    },
    body: JSON.stringify(jsonBody)
  }).then((response: SPHttpClientResponse) => {
    if (response.ok) {
      return response.json();
    }
  });

}

export function UpdateItem(WebUrl: string, spHttpClient: SPHttpClient, ListName: string, jsonBody: any, ID: number) {



  if (!jsonBody.__metadata) {
    jsonBody.__metadata = {
      'type': 'SP.ListItem'
    };
  }


  const URL = WebUrl + "/_api/web/lists/getbytitle('" + ListName + "')/Items(" + ID + ")";
  return spHttpClient.post(URL,
    SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify(jsonBody)
  }).then((response: SPHttpClientResponse) => {
    if (response.ok) {
      return response;
    }
  });

}


export async function getUserIdFromLoginName(context: WebPartContext, loginName: string): Promise<any> {
  const response = await context.spHttpClient.post(
    `${context.pageContext.web.absoluteUrl}/_api/web/ensureuser`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "odata-version": ""
      },
      body: JSON.stringify({ 'logonName': loginName })
    }
  );

  const userData = await response.json();
  return userData.d;
}





export async function UploadFile(WebUrl: string, spHttpClient: any, file: string, DisplayName: string | File, DocumentLib: string, jsonBody: { __metadata: { type: string; }; Name: string; TileLID: any; DocumentType: string; Documentpath: string; } | null): Promise<any> {

  // let fileupload = DocumentLib +"/"+FolderName;
  return new Promise((resolve) => {
    const spOpts: ISPHttpClientOptions = {
      body: file
    };
    const redirectionURL = WebUrl + "/_api/Web/GetFolderByServerRelativeUrl('" + DocumentLib + "')/Files/Add(url='" + DisplayName + "', overwrite=true)?$expand=ListItemAllFields";
    const responsedata = spHttpClient.post(redirectionURL, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {
      response.json().then(async (responseJSON: any) => {
        // console.log(responseJSON.ListItemAllFields.ID);
        const serverRelURL = await responseJSON.ServerRelativeUrl;
        if (jsonBody != null) {
          await UpdateItem(WebUrl, spHttpClient, DocumentLib, jsonBody, responseJSON.ListItemAllFields.ID);

        }
        resolve(responseJSON);
        console.log(responsedata);
        console.log(serverRelURL);
      });
    });
  });

}


export async function UpdateFileItem(context: WebPartContext, webURL: string, ListName: string, jsonBody: any, ID: number) {

  const URL = webURL + "/_api/web/lists/getbytitle('" + ListName + "')/Items(" + ID + ")";
  return await context.spHttpClient.post(URL,
    SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'MERGE'
    },
    body: JSON.stringify(jsonBody)
  }).then((response: SPHttpClientResponse) => {
    if (response.ok) {
      return response;
    }
  });

}
export async function uuidv4() {
  let tday = new Date();
  let d: any = tday.getDate();
  let m: any = tday.getMonth() + 1;
  let y = tday.getFullYear();
  let hr = tday.getHours();
  let min = tday.getMinutes();
  let sec = tday.getMilliseconds();
  if (d < 10) {
    d = '0' + d;
  }
  if (m < 10) {
    m = '0' + m;
  }

  let CreationDate = y + '-' + m + '-' + d + '-' + hr + '-' + min + '-' + sec;
  return CreationDate.toString();
}


/*export  function UploadFile(WebUrl,spHttpClient,file,DisplayName,DocumentLib,jsonBody,FolderName):Promise<any>  {
  
  let fileupload = DocumentLib +"/"+FolderName;
    return new Promise((resolve) => {      
        const spOpts: ISPHttpClientOptions = {      
            body: file      
        };      
        var redirectionURL = WebUrl + "/_api/Web/GetFolderByServerRelativeUrl('"+fileupload+"')/Files/Add(url='" + DisplayName + "', overwrite=true)?$expand=ListItemAllFields"      
        const response = spHttpClient.post(redirectionURL, SPHttpClient.configurations.v1, spOpts).then((response: SPHttpClientResponse) => {        
            response.json().then(async (responseJSON: any) => {        
               // console.log(responseJSON.ListItemAllFields.ID);
              var serverRelURL = await responseJSON.ServerRelativeUrl;    
               if(jsonBody != null){
                await UpdateItem(WebUrl,spHttpClient,DocumentLib,jsonBody,responseJSON.ListItemAllFields.ID)
            
               }
               resolve(responseJSON); 
            });        
          });    
        });   
    
}*/

export async function DeleteItem(WebUrl: string, spHttpClient: SPHttpClient, ListName: string, ID: number) {

  const URL = WebUrl + "/_api/web/lists/getbytitle('" + ListName + "')/Items(" + ID + ")";
  return await spHttpClient.post(URL,
    SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'odata-version': '3.0',
      'IF-MATCH': '*',
      'X-HTTP-Method': 'DELETE'
    }
  }).then((response: SPHttpClientResponse) => {
    if (response.ok) {
      return response;
    }
  });

}
function URLBuilder(url: string, options: any) {
  if (options) {
    if (options.filter) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$filter=" + options.filter;
    }
    if (options.select) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$select=" + options.select;
    }
    if (options.orderby) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$orderby=" + options.orderby;
    }
    if (options.expand) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$expand=" + options.expand;
    }
    if (options.top) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$top=" + options.top;
    }
    if (options.skip) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$skip=" + options.skip;
    }
    if (options.skiptoken) {
      url += ((url.indexOf('?') > -1) ? "&" : "?") + "$skiptoken=Paged%3DTRUE%26p_ID%3D" + options.skiptoken;
    }
  }
  return url;
}


// export async function FindUserGroup(context: WebPartContext, loginName: string): Promise<any> {
//   try {
//     const response: SPHttpClientResponse = await this.context.spHttpClient.get(
//       `${this.context.pageContext.web.absoluteUrl}/_api/Web/GetUserById(${this.context.pageContext.legacyPageContext.userId})?$expand=Groups`,
//       SPHttpClient.configurations.v1, {
//       headers: {
//         'Accept': 'application/json;odata=nometadata',
//         'odata-version': ''
//       }
//     }
//     );

//     if (response.ok) {
//       const data = await response.json();
//       userData(data.Groups); // Assuming Groups is the field that contains user groups
//     } else {
//       const errorMessage: string = `Error loading current user: ${response.status} - ${response.statusText}`;
//       console.log(new Error(errorMessage));
//     }
//   } catch (error) {
//     console.log(error instanceof Error ? error : new Error('Unknown error occurred'));
//   }
// }

// async function userData(groupData: any) {
//   let dinamicurl = "Permission/ID eq " + this.context.pageContext.legacyPageContext.userId;
//   try {
//     const response: SPHttpClientResponse = await this.context.spHttpClient.get(
//       `${this.context.pageContext.web.absoluteUrl}/_api/Web/siteusers`,
//       SPHttpClient.configurations.v1, {
//       headers: {
//         'Accept': 'application/json;odata=nometadata',
//         'odata-version': ''
//       }
//     }
//     );

//     if (response.ok) {
//       let data = await response.json();
//       let userArray = new Array();
//       data.value.map((el: any) => {
//         if (el.IsShareByEmailGuestUser == false) {
//           userArray.push(el);
//         }
//       });

//       var externaluser = userArray;
//       var NonExternalUser = externaluser.filter(Title => Title.Title == "Everyone except external users");
//       dinamicurl = dinamicurl + "or Permission/ID eq " + NonExternalUser[0].Id + " ";
//       for (var i = 0; i < groupData.length; i++) {
//         dinamicurl = dinamicurl + " or Permission/ID eq " + groupData[i].Id + " ";
//       }
//       this._loadCurrentUserDisplayName(dinamicurl);
//     } else {
//       const responseText: string = await response.text();
//       const errorMessage: string = `Error loading current user: ${response.status} - ${responseText}`;
//       console.log(new Error(errorMessage));
//     }
//   } catch (error) {
//     console.log(error instanceof Error ? error : new Error(error));
//   }
// }

export async function isMember(context: WebPartContext, GroupName: string) {
  let url = context.pageContext.web.absoluteUrl + "/_api/web/sitegroups/getByName('" + GroupName + "')/Users?$filter=Id eq " + context.pageContext.legacyPageContext["userId"];
  return await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': ''
    },
  }).then((response: SPHttpClientResponse) => {
    return response.json();
  });
}





