import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode, Selection } from '@fluentui/react/lib/DetailsList';
import { DefaultButton } from '@fluentui/react';
import { FilesService } from '../../../services/SharePointService';
import { IFilesProps, IFileItem } from './IFilesProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { FileUpload } from '../FileUpload/FileUpload';
import { ModalContainer } from '../ModalContainer/ModalContainer';
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
    const handleDeleteFiles = async () => {
        if (selectedFiles.length === 0) return;
        setShowDeleteConfirm(true);
    };

    const handleConfirmDelete = async () => {
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
        } catch (error) {
            alert('Errore durante l\'eliminazione dei file');
        } finally {
            setSaving(false);
            setUploading(false);
        }
    };

    const handleFileUpload = async (file: File) => {
        setUploading(true);
        try {
            const service = new FilesService(context);
            await service.uploadFile(file);
            // Refresh list after upload
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
        } catch (error) {
            alert('Errore durante il caricamento del file');
        } finally {
            setUploading(false);
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
            }
            catch {
                setItems([]);
            }
            finally {
                setLoading(false);
            }
        };
        void fetchFiles();
    }
        , []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    }
    const columns: IColumn[] = [
        { key: 'FileName', name: 'File Name', fieldName: 'FileName', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'CheckedOutTo', name: 'Checked Out To', fieldName: 'CheckedOutTo', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.CheckedOutTo ? item.CheckedOutTo.Title : '' },
        { key: 'CreatedBy', name: 'Created By', fieldName: 'CreatedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.CreatedBy || '' },
        { key: 'Created', name: 'Created', fieldName: 'Created', onRender: (item) => formatDate(item.Created), minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'ModifiedBy', name: 'Modified By', fieldName: 'ModifiedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.ModifiedBy || '' },
        { key: 'Modified', name: 'Modified', fieldName: 'Modified', onRender: (item) => formatDate(item.Modified), minWidth: 100, maxWidth: 140, isResizable: true },
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
