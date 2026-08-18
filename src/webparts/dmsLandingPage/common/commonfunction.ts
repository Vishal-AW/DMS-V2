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
            // id: folder.Id,
            id: folder.ID,
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

    const sortChildren = (nodes: any[]) => {
        nodes.sort((a: any, b: any) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
        nodes.forEach(node => {
            if (node.children.length > 0) {
                sortChildren(node.children);
            }
        });
    };

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
    sortChildren(tree);
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

export const getChildFolders = async (
    context: WebPartContext,
    folderPath: string
) => {
    const sp = spfi().using(SPFx(context));

    try {
        const subfolders = await sp.web
            .getFolderByServerRelativePath(folderPath)
            .folders
            .select("Name", "ServerRelativeUrl", "ItemCount", "ListItemAllFields/*")
            .expand("ListItemAllFields")();

        const nodes = subfolders
            .filter((f: any) => f.Name && f.Name.toLowerCase() !== "forms" && !f.Name.startsWith("."))
            .map((f: any) => ({
                id: f.ListItemAllFields?.ID ? String(f.ListItemAllFields.ID) : f.ServerRelativeUrl,
                name: f.Name,
                path: f.ServerRelativeUrl,
                children: [],
                isLoaded: false,
                isLoading: false,
                isLastLevel: false,
                FileRef: f.ServerRelativeUrl,
                FileDirRef: folderPath,
                FSObjType: 1,
                ItemCount: f.ItemCount,
                Modified: f.ListItemAllFields?.Modified,
                Created: f.ListItemAllFields?.Created,
                Editor: f.ListItemAllFields?.Editor,
                Author: f.ListItemAllFields?.Author
            }))
            .sort((a: any, b: any) => a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));

        // SharePoint's ListItemAllFields projection on the /Folders endpoint does
        // not always include the display names of the person fields (Author/Editor).
        // Enrich them from the library's list items using the same
        // Author/Title + expand("Author,Editor") pattern used across the app.
        // This step is optional: if it fails, the folder list still works exactly
        // as before and the grid simply keeps the (possibly blank) person columns.
        try {
            const listId = (subfolders[0] as any)?.ListItemAllFields?.ListId;
            // Fallback: when ListItemAllFields does not expose ListId, derive the
            // library name from the folder path (libraries live at the web root).
            const webRelativeUrl = context.pageContext.web.serverRelativeUrl;
            const libraryName = folderPath.replace(webRelativeUrl, "").split("/").filter(Boolean)[0];

            let folderItems: any[] = [];
            if (listId && nodes.length > 0) {
                folderItems = await sp.web.lists.getById(listId).items
                    .select("Id", "FileRef", "Author/Title", "Editor/Title")
                    .expand("Author,Editor")
                    .filter(`FSObjType eq 1 and FileDirRef eq '${folderPath}'`)
                    .top(5000)();
            } else if (libraryName && nodes.length > 0) {
                folderItems = await sp.web.lists.getByTitle(libraryName).items
                    .select("Id", "FileRef", "Author/Title", "Editor/Title")
                    .expand("Author,Editor")
                    .filter(`FSObjType eq 1 and FileDirRef eq '${folderPath}'`)
                    .top(5000)();
            }
            const itemsByPath = new Map<string, any>(folderItems.map((it: any) => [it.FileRef, it]));
            nodes.forEach((node: any) => {
                const item = itemsByPath.get(node.path);
                if (item) {
                    if (item.Author?.Title) node.Author = { Id: item.Author.Id, Title: item.Author.Title };
                    if (item.Editor?.Title) node.Editor = { Id: item.Editor.Id, Title: item.Editor.Title };
                }
            });

            // TEMP-DIAG: verify person metadata on the grid rows. Remove after confirming.
            if (nodes.length > 0) {
                console.log("[DMS-TEMP] child folders enriched:", {
                    folderPath,
                    count: nodes.length,
                    sample: {
                        name: nodes[0].name,
                        Created: nodes[0].Created,
                        Modified: nodes[0].Modified,
                        Author: nodes[0].Author,
                        Editor: nodes[0].Editor
                    }
                });
            }
        } catch (enrichError) {
            console.warn("Unable to enrich folder author/editor names for path:", folderPath, enrichError);
        }

        return nodes;
    } catch (error) {
        console.error("Error fetching child folders for path:", folderPath, error);
        return [];
    }
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
