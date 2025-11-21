import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { FilesService } from '../../../services/SharePointService';
import { IFilesProps, IFileItem } from './IFilesProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
export const Files: React.FC<IFilesProps> = ({ context }) => {
    const [items, setItems] = React.useState<IFileItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchFiles = async (): Promise<void> => {
            try {
                const service = new FilesService(context);
                const data = await service.getFiles();
                const mapped: IFileItem[] = data.map((item) => ({
                    FileName: item.FileLeafRef || '',
                    CheckedOutTo: item.CheckoutUser ? { Id: item.CheckoutUser.Id, Title: item.CheckoutUser.Title } : undefined,
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author.Title,
                    ModifiedBy: item.Editor.Title,
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
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    <DetailsList
                        items={items}
                        columns={columns}
                        setKey="set"
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        constrainMode={ConstrainMode.unconstrained}
                    />
                </div>
            )}
        </div>
    );
};
