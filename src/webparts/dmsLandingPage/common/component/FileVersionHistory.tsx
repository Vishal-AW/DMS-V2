import * as React from 'react';
import { DetailsList, IColumn, SelectionMode, PrimaryButton, Dialog, DialogType } from '@fluentui/react';
import { getFileVersionsByUrl, IFileVersion } from '../../../../Services/FileVersionService';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IFileVersionHistoryProps {
    context: WebPartContext;
    fileServerRelativeUrl: string;   // e.g., "/sites/MySite/Shared Documents/MyFile.docx"
    fileName: string;
}

const FileVersionHistory: React.FC<IFileVersionHistoryProps> = (props) => {
    const [versions, setVersions] = React.useState<IFileVersion[]>([]);
    const [isLoading, setIsLoading] = React.useState<boolean>(false);
    const [selectedVersion, setSelectedVersion] = React.useState<any>(null);
    const [isDialogOpen, setIsDialogOpen] = React.useState<boolean>(false);

    const loadVersions = async () => {
        setIsLoading(true);
        try {
            const data = await getFileVersionsByUrl(props.context, props.fileServerRelativeUrl);
            setVersions(data);
        } catch (err) {
            alert("Failed to load versions");
        }
        setIsLoading(false);
    };

    React.useEffect(() => {
        if (props.fileServerRelativeUrl) {
            loadVersions();
        }
    }, [props.fileServerRelativeUrl]);

    const columns: IColumn[] = [
        { key: 'version', name: 'Version', fieldName: 'VersionLabel', minWidth: 80, maxWidth: 100 },
        { key: 'modified', name: 'Modified', fieldName: 'Modified', minWidth: 150, isResizable: true },
        { key: 'editor', name: 'Modified By', fieldName: 'Editor.Title', minWidth: 150 },
        {
            key: 'size', name: 'Size (KB)', fieldName: 'Size', minWidth: 100,
            onRender: (item) => (item.Size / 1024).toFixed(2) + ' KB'
        },
        { key: 'comment', name: 'Comment', fieldName: 'CheckInComment', minWidth: 200 },
        {
            key: 'actions',
            name: 'Actions',
            minWidth: 150,
            onRender: (item) => (
                <>
                    <PrimaryButton
                        text="View"
                        onClick={() => { setSelectedVersion(item); setIsDialogOpen(true); }}
                        style={{ marginRight: 8 }}
                    />
                    <PrimaryButton
                        text="Restore"
                    // onClick={() => service.restoreVersion(props.fileServerRelativeUrl, item.VersionLabel)}
                    />
                </>
            )
        }
    ];

    return (
        <div>
            <h3>Version History - {props.fileName}</h3>

            <DetailsList
                items={versions}
                columns={columns}
                selectionMode={SelectionMode.none}
                isHeaderVisible={true}
                onItemInvoked={(item) => { setSelectedVersion(item); setIsDialogOpen(true); }}
            />

            <Dialog
                hidden={!isDialogOpen}
                onDismiss={() => setIsDialogOpen(false)}
                dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: `Version ${selectedVersion?.VersionLabel}`,
                }}
                minWidth={600}
            >
                <pre>{JSON.stringify(selectedVersion, null, 2)}</pre>
            </Dialog>
        </div>
    );
};

export default FileVersionHistory;