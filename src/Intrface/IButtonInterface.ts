export interface IButtonsProps {
    Id: number;
    ButtonDisplayName: string;
    ButtonType: string;
    Title: string;
    key: string;
    value: boolean;
    InternalName: string;
    Active: boolean;
    Sequence?: number;
    Icons: string;
}
// interface IUserId {
//     Id: number;
//     Title: string;
// }
export interface IRolePermission {
    Role: string;
    UsersId: any[];
    Permission: IButtonsProps[];
}