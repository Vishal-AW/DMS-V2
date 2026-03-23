export interface IColumnSchema {
    name: string;
    ColType: string;
    LookupList?: string;
    LookupField?: string;
    DefaultValue?: string;
    choices?: string[];
}

export interface IListSchema {
    title: string;
    description?: string;
    templateType: number;
    columns: IColumnSchema[];
}
