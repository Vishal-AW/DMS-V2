/* eslint-disable */
import { GetListItem, UpdateItem } from '../DAL/Commonfile';
import { SPHttpClient } from '@microsoft/sp-http';

export function getAllButtonsAdmin(WebUrl: string, spHttpClient: SPHttpClient) {
  const option = {
    select: 'ID,Title,InternalName,Active,Sequence,ButtonType,ButtonDisplayName,Icons,FullControl,Contribute,Edit,Read',
    orderby: 'Sequence asc',
    top: 5000
  };
  return GetListItem(WebUrl, spHttpClient, 'DMS_Buttons', option);
}

export function updateButtonItem(
  WebUrl: string,
  spHttpClient: SPHttpClient,
  id: number,
  data: {
    Title?: string;
    InternalName?: string;
    Active?: boolean;
    Sequence?: number;
    ButtonType?: string;
    ButtonDisplayName?: string;
    Icons?: string;
    FullControl?: boolean;
    Contribute?: boolean;
    Edit?: boolean;
    Read?: boolean;
  }
) {
  return UpdateItem(WebUrl, spHttpClient, 'DMS_Buttons', data, id);
}
