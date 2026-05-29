export interface IUserPermissions {
    hasFullControl: boolean;
    hasContribute: boolean;
    hasEdit: boolean;
    hasRead: boolean;
}

export interface IDmsButton {
    Title: string;
    InternalName: string;
    ButtonType: string;
    ButtonDisplayName: string;
    Icons: string;
    Sequence: number;
    FullControl: boolean;
    Contribute: boolean;
    EditPermission: boolean;
    ReadPermission: boolean;
}

export function isButtonPermitted(
    btn: IDmsButton,
    perms: IUserPermissions
): boolean {
    if (perms.hasFullControl) return true;

    // Hierarchical check: A user with Edit can see Edit, Contribute, and Read buttons.
    if (perms.hasEdit) return btn.EditPermission || btn.Contribute || btn.ReadPermission;
    if (perms.hasContribute) return btn.Contribute || btn.ReadPermission;
    if (perms.hasRead) return btn.ReadPermission;

    return false;
}