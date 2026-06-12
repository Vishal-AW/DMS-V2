export interface IColumnSchema {
    name: string;
    ColType: string;
    LookupList?: string;
    LookupField?: string;
    DefaultValue?: string;
    choices?: string[];
    indexed?: boolean;   //used for indexing here
}

export interface IListSchema {
    title: string;
    description?: string;
    templateType: number;
    columns: IColumnSchema[];
}
