import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";

import { DocumentPdf20Regular, Document20Regular, DocumentTable20Regular, Image20Regular, Cube20Regular, DocumentText20Regular } from '@fluentui/react-icons';

export function buildBreadcrumbPath(folder: any, allFolders: any[]): any[] {
    const path: any[] = [];

    function findPath(nodes: any[], target: string): boolean {
        for (const node of nodes) {
            if (node.id === target) {
                path.push(node);
                return true;
            }
            if (node.children) {
                if (findPath(node.children, target)) {
                    path.unshift(node);
                    return true;
                }
            }
        }
        return false;
    }

    findPath(allFolders, folder.id);
    return path;
}
export const buildFolderHierarchy = (
    folders: any[],
    libraryRoot: string
): any[] => {

    const map = new Map<string, any>();
    const tree: any[] = [];
    folders.forEach(folder => {
        const name = folder.FileRef.split("/").pop() || "";

        map.set(folder.FileRef, {
            id: folder.Id,
            name,
            path: folder.FileRef,
            children: [],
            isLastLevel: false,
            ...folder,
        });
    });

    folders.forEach(folder => {
        const node = map.get(folder.FileRef)!;
        if (folder.FileDirRef === libraryRoot) {
            tree.push(node);
        } else {
            const parent = map.get(folder.FileDirRef);
            if (parent) {
                parent.children.push(node);
            }
        }
    });

    const markLastLevel = (nodes: any[]) => {
        nodes.forEach(node => {
            if (node.children.length === 0) {
                node.isLastLevel = true;
            } else {
                markLastLevel(node.children);
            }
        });
    };

    markLastLevel(tree);
    return tree;
};

export const buildLibraryRootPath = (context: WebPartContext, libName: string) => {
    const webRelativeUrl = context.pageContext.web.serverRelativeUrl;

    return webRelativeUrl === "/"
        ? `/${libName}`
        : `${webRelativeUrl}/${libName}`;
};

export const getAllDocuments = async (
    context: WebPartContext,
    folderPath: string
) => {
    const sp = spfi().using(SPFx(context));

    const files = await sp.web
        .getFolderByServerRelativePath(folderPath)
        .files
        .select("*,ListItemAllFields/*,CheckedOutByUser")
        .expand("ListItemAllFields,CheckedOutByUser")();

    return files;
};

export const fileTypeConfig: Record<string, { IconName: typeof DocumentPdf20Regular; className: string; label: string; }> = {
    pdf: { IconName: DocumentPdf20Regular, className: 'doc-icon-pdf', label: 'PDF' },
    docx: { IconName: DocumentText20Regular, className: 'doc-icon-word', label: 'Word' },
    xlsx: { IconName: DocumentTable20Regular, className: 'doc-icon-excel', label: 'Excel' },
    png: { IconName: Image20Regular, className: 'doc-icon-image', label: 'Image' },
    jpg: { IconName: Image20Regular, className: 'doc-icon-image', label: 'Image' },
    dwg: { IconName: Cube20Regular, className: 'doc-icon-cad', label: 'AutoCAD' },
    other: { IconName: Document20Regular, className: 'doc-icon-other', label: 'File' },
};

export const checkExtension = (fileName: string): boolean => {
    if (!fileName) return false;
    const extension = fileName.split(".").pop()?.toLowerCase();
    const allowedExtensions = ["pdf", "txt", "jpg", "jpeg", "png", "gif", "bmp"];
    return !allowedExtensions.includes(extension || "");
};

export const checkButtons = (input: string): boolean => {
    if (!input) return false;
    const buttonTypes = ["OpenInApp", "CheckIn", "DiscardCheckOut", "CheckOut", "Preview"];
    return !buttonTypes.includes(input);
};

export const getOpenAppURL = (filePath: string, SiteURL: string) => {
    const portalUrl = new URL(SiteURL).origin;
    if (!filePath) return;
    const extension = filePath.split('.').pop()?.toLowerCase();
    if (!extension) return;

    let appUrl: string | null = null;
    switch (extension) {
        case 'xls':
        case 'xlsx':
            appUrl = `ms-excel:ofe|u|${portalUrl}${filePath}`;
            break;
        case 'doc':
        case 'docx':
            appUrl = `ms-word:ofe|u|${portalUrl}${filePath}`;
            break;
        case 'ppt':
        case 'pptx':
            appUrl = `ms-powerpoint:ofe|u|${portalUrl}${filePath}`;
            break;
    }

    if (appUrl) {
        window.open(appUrl, '_blank');
    }
};