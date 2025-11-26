import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, Selection } from '@fluentui/react/lib/DetailsList';
import { DefaultButton } from '@fluentui/react';
import { FilesService, DocumentsService } from '../../../services/SharePointService';
import { IFilesProps, IFileItem } from './IFilesProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { FileUpload } from '../FileUpload/FileUpload';
import { DocumentMetadataForm } from './DocumentMetadataForm';
import { IDocumentMetadataProps } from './IDocumentMetadataFormProps';
import { ModalContainer } from '../ModalContainer/ModalContainer';
import { UserHoverCardSmart } from '../UserHoverCard/UserHoverCard';

export const Files: React.FC<IFilesProps> = ({ context }) => {
    const [selectedFiles, setSelectedFiles] = React.useState<IFileItem[]>([]);
    const selection = React.useRef(new Selection({
        onSelectionChanged: () => {
            setSelectedFiles(selection.current.getSelection() as IFileItem[]);
        }
    }));
    const [uploading, setUploading] = React.useState(false);
    const [showDeleteConfirm, setShowDeleteConfirm] = React.useState(false);
    const [saving, setSaving] = React.useState(false);
    const [items, setItems] = React.useState<IFileItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [showMetadataModal, setShowMetadataModal] = React.useState(false);
    const [pendingFile, setPendingFile] = React.useState<File | null>(null);

    const handleDeleteFiles = async (): Promise<void> => {
        if (selectedFiles.length === 0) return;
        setShowDeleteConfirm(true);
    }

    const handleConfirmDelete = async (): Promise<void> => {
        setSaving(true);
        try {
            const service = new FilesService(context);
            // Filter out undefined Ids to ensure type safety
            const ids = selectedFiles.map(f => f.Id).filter((id): id is number => id !== undefined);
            await service.deleteFilesById(ids);
            // Refresh list after delete
            const data = await service.getFiles();
            const mapped: IFileItem[] = data.map((item) => ({
                Id: item.Id,
                FileName: item.FileLeafRef || '',
                CheckedOutTo: item.CheckoutUser ? { Id: item.CheckoutUser.Id, Title: item.CheckoutUser.Title } : undefined,
                Modified: item.Modified || '',
                Created: item.Created || '',
                CreatedBy: item.Author.Title,
                ModifiedBy: item.Editor.Title,
            }));
            setItems(mapped);
            setSelectedFiles([]);
            selection.current.setAllSelected(false);
            setShowDeleteConfirm(false);
        } catch {
            alert('Errore durante l\'eliminazione dei file');
        } finally {
            setSaving(false);
            setUploading(false);
        }
    };

    // Step 1: User selects file, show metadata modal
    const handleFileUpload = (file: File): void => {
        setPendingFile(file);
        setShowMetadataModal(true);
    };

    // Step 2: User submits metadata, upload file and create Document
    const handleMetadataSubmit = async (metadata: IDocumentMetadataProps): Promise<void> => {
        if (!pendingFile) return;
        setShowMetadataModal(false);
        setUploading(true);
        try {
            const filesService = new FilesService(context);
            await filesService.uploadFile(pendingFile);
            const documentsService = new DocumentsService(context);
            await documentsService.createDocument(metadata);
            // Refresh list after upload
            const data = await filesService.getFiles();
            const mapped: IFileItem[] = data.map((item) => ({
                Id: item.Id,
                FileName: item.FileLeafRef || '',
                CheckedOutTo: item.CheckoutUser ? { Id: item.CheckoutUser.Id, Title: item.CheckoutUser.Title } : undefined,
                Modified: item.Modified || '',
                Created: item.Created || '',
                CreatedBy: item.Author?.Title,
                ModifiedBy: item.Editor?.Title,
            }));
            setItems(mapped);
        } catch {
            alert('Error while uploading the file and creating the Document record');
        } finally {
            setUploading(false);
            setUploading(false);
            setPendingFile(null);
        }
    };

    React.useEffect(() => {
        const fetchFiles = async (): Promise<void> => {
            try {
                const service = new FilesService(context);
                const data = await service.getFiles();
                const mapped: IFileItem[] = data.map((item) => ({
                    Id: item.Id,
                    FileName: item.FileLeafRef || '',
                    CheckedOutTo: item.CheckoutUser ? { Id: item.CheckoutUser.Id, Title: item.CheckoutUser.Title } : undefined,
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author?.Title,
                    ModifiedBy: item.Editor?.Title,
                }));
                setItems(mapped);
                console.log('Fetched files:', mapped);
            } catch {
                // Optionally handle error, e.g. set error state
            } finally {
                setLoading(false);
            }
        };
        void fetchFiles();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    // Define columns outside of render to avoid re-creation on each render
    const columns: IColumn[] = [
        {
            key: 'FileName',
            name: 'FileName',
            fieldName: 'FileName',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
        },
        {
            key: 'CreatedBy',
            name: 'CreatedBy',
            fieldName: 'CreatedBy',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
            onRender: (item) => {
                if (!item.CreatedBy || !item.CreatedBy.EMail) return '';
                return (
                    <UserHoverCardSmart
                        email={item.CreatedBy.EMail}
                        displayName={item.CreatedBy.Title}
                        pictureUrl={item.CreatedBy.Picture}
                        context={context}
                    />
                );
            }
        },
        {
            key: 'Created',
            name: 'Created',
            fieldName: 'Created',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
        },
        {
            key: 'ModifiedBy',
            name: 'ModifiedBy',
            fieldName: 'ModifiedBy',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
            onRender: (item) => {
                if (!item.ModifiedBy || !item.ModifiedBy.EMail) return '';
                return (
                    <UserHoverCardSmart
                        email={item.ModifiedBy.EMail}
                        displayName={item.ModifiedBy.Title}
                        pictureUrl={item.ModifiedBy.Picture}
                        context={context}
                    />
                );
            }
        },
        {
            key: 'Modified',
            name: 'Modified',
            fieldName: 'Modified',
            minWidth: 100,
            maxWidth: 200,
            isResizable: true,
        },
    ];

    return (
        <div className={styles.container}>
            <FileUpload onUpload={handleFileUpload} disabled={uploading || loading} />
            <DefaultButton
                text="Elimina selezionati"
                onClick={handleDeleteFiles}
                disabled={uploading || loading || selectedFiles.length === 0}
                style={{ marginBottom: 12 }}
            />
            {showMetadataModal && (
                <ModalContainer
                    isOpen={showMetadataModal}
                    title="Document Metadata"
                    onCancel={() => { setShowMetadataModal(false); setPendingFile(null); }}
                >
                    <DocumentMetadataForm
                        context={context}
                        initialValues={pendingFile ? { DocumentCode: pendingFile.name, Title: pendingFile.name } : {}}
                        onSubmit={handleMetadataSubmit}
                        onCancel={() => { setShowMetadataModal(false); setPendingFile(null); }}
                        saving={uploading}
                    />
                </ModalContainer>
            )}
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    <DetailsList
                        items={items}
                        columns={columns}
                        setKey="multiple"
                        selection={selection.current}
                        selectionMode={2}
                        selectionPreservedOnEmptyClick={true}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        constrainMode={ConstrainMode.horizontalConstrained}
                        isHeaderVisible={true}
                        ariaLabelForSelectionColumn="Toggle selection"
                        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                        checkButtonAriaLabel="select row"
                    />
                </div>
            )}
            {showDeleteConfirm && (
                <ModalContainer
                    isOpen={showDeleteConfirm}
                    title={selectedFiles.length === 1 ? 'Conferma eliminazione file' : `Conferma eliminazione di ${selectedFiles.length} file`}
                    onSave={handleConfirmDelete}
                    onCancel={() => setShowDeleteConfirm(false)}
                    saving={saving}
                    saveText="Elimina"
                    cancelText="Annulla"
                >
                    <div className={styles.deleteConfirmText}>
                        Sei sicuro di voler eliminare {selectedFiles.length === 1 ? 'questo file' : `i ${selectedFiles.length} file selezionati`}?
                    </div>
                </ModalContainer>
            )}
        </div>
    );
};
