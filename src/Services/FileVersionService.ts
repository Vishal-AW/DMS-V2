import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/files";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFileVersion {
    VersionLabel: string;
    Created: string;
    Modified: string;
    Editor?: { Title: string; Email?: string; };
    Size?: number;
    CheckInComment?: string;
}

export const getFileVersionsByUrl = async (context: WebPartContext, serverRelativeUrl: string): Promise<IFileVersion[]> => {
    try {
        const sp: SPFI = spfi().using(SPFx(context));
        const versions = await sp.web
            .getFileByServerRelativePath(serverRelativeUrl)   // ← This is the correct method now
            .versions
            .select("VersionLabel", "Created", "Modified", "Size", "CheckInComment", "Editor/Title", "Editor/Email")
            .expand("Editor")();

        return versions;
    } catch (error) {
        console.error("Error fetching file versions:", error);
        throw error;
    }
};