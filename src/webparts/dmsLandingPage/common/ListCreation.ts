/* eslint-disable */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IColumnSchema, IListSchema } from "../../../Intrface/IListSchema";
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import '@pnp/sp/fields';
import '@pnp/sp/site-users/web';
import '@pnp/sp/attachments';
import '@pnp/sp/views';
import { IList } from "@pnp/sp/lists";
import { FieldUserSelectionMode } from "@pnp/sp/fields";
import { SPHttpClient } from "@microsoft/sp-http-base";
import { UpdateTileSetting } from "../../../Services/MasTileService";

export const TileLibrary = async (context: WebPartContext, Internal: string, TileLID: number, ArchiveInternal: string, isUpdate: boolean, tableData: any, IsArchiveAllowed: boolean,) => {
    const Columns: IListSchema[] = [{
        title: isUpdate ? ArchiveInternal : Internal,
        templateType: 101,
        columns: [
            { "name": "DefineRole", "ColType": "8" },
            { "name": "ProjectmanagerAllow", "ColType": "8" },
            { "name": "Projectmanager", "ColType": "20" },
            { "name": "ProjectmanagerEmail", "ColType": "2" },
            { "name": "PublisherAllow", "ColType": "8" },
            { "name": "Publisher", "ColType": "20" },
            { "name": "PublisherEmail", "ColType": "2" },
            { "name": "CurrentApprover", "ColType": "2" },
            { "name": "Status", "ColType": "7", "LookupField": "StatusName", "LookupList": "DMS_Mas_Status" },
            { "name": "InternalStatus", "ColType": "2" },
            { "name": "ProjectMasterLID", "ColType": "2" },
            { "name": "LatestRemark", "ColType": "3" },
            { "name": "AllowApprover", "ColType": "8" },
            { "name": "Active", "ColType": "8" },
            { "name": "DisplayStatus", "ColType": "2" },
            { "name": "ReferenceNo", "ColType": "2" },
            { "name": "RefSequence", "ColType": "9" },
            { "name": "Level", "ColType": "2" },
            { "name": "Revision", "ColType": "2" },
            { "name": "DocStatus", "ColType": "2" },
            { "name": "Template", "ColType": "2" },
            { "name": "CreateFolder", "ColType": "8" },
            { "name": "Company", "ColType": "2" },
            { "name": "ActualName", "ColType": "2" },
            { "name": "DocumentSuffix", "ColType": "2" },
            { "name": "OtherSuffix", "ColType": "2" },
            { "name": "PSType", "ColType": "2" },
            { "name": "IsArchiveFlag", "ColType": "8" },
            { "name": "IsExistingRefID", "ColType": "9" },
            { "name": "IsExistingFlag", "ColType": "2" },
            { "name": "OCRText", "ColType": "3" },
            { "name": "DeleteFlag", "ColType": "2" },
            { "name": "OCRText0", "ColType": "3" },
            { "name": "OCRText1", "ColType": "3" },
            { "name": "OCRText2", "ColType": "3" },
            { "name": "OCRText3", "ColType": "3" },
            { "name": "OCRText4", "ColType": "3" },
            { "name": "OCRText5", "ColType": "3" },
            { "name": "OCRText6", "ColType": "3" },
            { "name": "OCRText7", "ColType": "3" },
            { "name": "OCRText8", "ColType": "3" },
            { "name": "OCRText9", "ColType": "3" },
            { "name": "IsSuffixRequired", "ColType": "8" },
            { "name": "FolderDocumentPath", "ColType": "3" },
            { "name": "OCRStatus", "ColType": "2" },
            { "name": "UploadFlag", "ColType": "2", "DefaultValue": "Backend" },
            { "name": "NewFolderAccess", "ColType": "2" },
        ]
    }];

    if (tableData.length > 0) {
        tableData.map(function (el: any) {
            let colType = getColumnType(el.ColumnType);
            Columns[0].columns.push({ "name": el.InternalTitleName, "ColType": colType.toString() });
        });
    }

    if (IsArchiveAllowed && !isUpdate) {
        const obj: IListSchema = {
            title: ArchiveInternal,
            templateType: 101,
            columns: Columns[0].columns

        };
        Columns.push(obj);
    }

    await createList(context, Columns, TileLID, false);

};

export const getColumnType = (val: any) => {
    switch (val) {
        case 'Multiple lines of Text':
            return 3;

        case 'Date and Time':
            return 4;

        case 'Choice':
            return 6;

        case 'Lookup':
            return 7;

        case 'Yes/No':
            return 8;

        case 'Number':
            return 9;

        case 'Person or Group':
            return 20;

        default:
            return 2;
    }
};


export const createColumn = async (
    list: IList,
    col: IColumnSchema,
    context: WebPartContext
) => {
    try {
        const sp = spfi().using(SPFx(context));
        switch (col.ColType) {

            case "2":
                await list.fields.addText(col.name, {
                    MaxLength: 255,
                });
                break;

            case "6":
                await list.fields.addChoice(col.name, {
                    Choices: col.choices || [],
                });
                break;

            case "9":
                await list.fields.addNumber(col.name);
                break;

            case "4":
                await list.fields.addDateTime(col.name);
                break;

            case "8":
                await list.fields.addBoolean(col.name);
                break;

            case "20":
                await list.fields.addUser(col.name, {
                    SelectionMode: FieldUserSelectionMode.PeopleAndGroups
                });
                break;

            case "7":

                let LookUplist;
                try {
                    LookUplist = await sp.web.lists.getByTitle(col.LookupList!).select("Id")();
                } catch {
                    const created = await sp.web.lists.add(
                        col.LookupList!
                    );
                    LookUplist = created;
                }

                await list.fields.addLookup(col.name, {
                    LookupListId: LookUplist.Id,
                    LookupFieldName: col.LookupField!

                });
                break;
            case "3":
                await list.fields.addMultilineText(col.name, {
                    NumberOfLines: 6,
                    RichText: false,
                    RestrictedMode: false,
                    AppendOnly: false
                });
                break;
        }
    } catch (error) {
        console.error(`Error creating column ${col.name}: `, error);
    }

};


export const addFieldsToDefaultView = async (
    list: IList,
    fields: string[]
) => {
    const view = await list.defaultView();


    for (const field of fields) {
        try {
            await list.views.getById(view.Id).fields.add(field);
        } catch (error) {
            // ignore duplicate field errors
            console.warn(`Error adding field ${field} to view: ${error}`);
        }
    }
};


const createList = async (context: WebPartContext, listData: IListSchema[], TileLID: number, isArchive: boolean) => {

    const sp = spfi().using(SPFx(context));
    for (const listDef of listData) {
        const listEnsureResult = await sp.web.lists.ensure(listDef.title, "Project Documents Library", listDef.templateType);
        const list = listEnsureResult.list;
        const listData = await list.select("Id")();
        const obj = { LibGuidName: listData?.Id };
        isArchive ? "" : await UpdateTileSetting(context.pageContext.web.absoluteUrl, context.spHttpClient, obj, TileLID);
        const viewFields: string[] = [];
        for (const col of listDef.columns) {
            await createColumn(list, col, context);
            viewFields.push(col.name);
        }
        await addFieldsToDefaultView(list, viewFields);
    }
};




export async function GetListData(context: WebPartContext, query: string) {
    const response = await context.spHttpClient.get(query, SPHttpClient.configurations.v1, {
        headers: {
            'Accept': 'application/json;odata=verbose',
            'odata-version': '',
        },
    });
    return await response.json();
};