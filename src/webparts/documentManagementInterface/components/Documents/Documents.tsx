import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { DocumentsService } from '../../../services/SharePointService';
import { IDocumentsProps, IDocumentItem } from './IDocumentsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';

export const Documents: React.FC<IDocumentsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDocumentItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchDocuments = async (): Promise<void> => {
            try {
                const service = new DocumentsService(context);
                const data = await service.getDocuments();
                const stripHtml = (html: string) => html.replace(/<[^>]+>/g, '').trim();
                const mapped: IDocumentItem[] = data.map((item) => ({
                    DocumentCode: item.DocumentCode || '',
                    Title: item.Title || '',
                    Revision: item.Revision || '',
                    Status: item.Status || '',
                    IssuePurpose: item.IssuePurpose || '',
                    ApprovalCode: item.ApprovalCode || '',
                    SentDate: item.SentDate || '',
                    ExpectedReturnDate: item.ExpectedReturnDate || '',
                    ActualReturnDate: item.ActualReturnDate || '',
                    TurnaroundDays: item.TurnaroundDays || 0,
                    DaysLate: item.DaysLate || 0,
                    AssignedTo: item.AssignedTo ? { Id: item.AssignedTo.Id, Title: item.AssignedTo.Title } : undefined,
                    Notes: item.Notes ? stripHtml(item.Notes) : '',
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author?.Title || '',
                    ModifiedBy: item.Editor?.Title || '',
                }));
                setItems(mapped);
                console.log('Fetched documents:', mapped);
            }
            catch {
                setItems([]);
            }
            finally {
                setLoading(false);
            }
        };
        void fetchDocuments();
    }, []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };

    const columns: IColumn[] = [
        { key: 'DocumentCode', name: 'Document Code', fieldName: 'DocumentCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Revision', name: 'Revision', fieldName: 'Revision', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'IssuePurpose', name: 'Issue Purpose', fieldName: 'IssuePurpose', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'ApprovalCode', name: 'Approval Code', fieldName: 'ApprovalCode', minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'SentDate', name: 'Sent Date', fieldName: 'SentDate', minWidth: 100, maxWidth: 140, isResizable: true, onRender: (item) => formatDate(item.SentDate) },
        { key: 'ExpectedReturnDate', name: 'Expected Return Date', fieldName: 'ExpectedReturnDate', minWidth: 140, maxWidth: 180, isResizable: true, onRender: (item) => formatDate(item.ExpectedReturnDate) },
        { key: 'ActualReturnDate', name: 'Actual Return Date', fieldName: 'ActualReturnDate', minWidth: 120, maxWidth: 160, isResizable: true, onRender: (item) => formatDate(item.ActualReturnDate) },
        { key: 'TurnaroundDays', name: 'Turnaround Days', fieldName: 'TurnaroundDays', minWidth: 120, maxWidth: 160, isResizable: true, onRender: (item) => Math.floor(item.TurnaroundDays) },
        { key: 'DaysLate', name: 'Days Late', fieldName: 'DaysLate', minWidth: 80, maxWidth: 120, isResizable: true, onRender: (item) => Math.floor(item.DaysLate) },
        { key: 'AssignedTo', name: 'Assigned To', fieldName: 'AssignedTo', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.AssignedTo?.Title || '' },
        { key: 'Notes', name: 'Notes', fieldName: 'Notes', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'CreatedBy', name: 'Created By', fieldName: 'CreatedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.CreatedBy || '' },
        { key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 100, maxWidth: 140, isResizable: true, onRender: (item) => formatDate(item.Created) },
        { key: 'ModifiedBy', name: 'Modified By', fieldName: 'ModifiedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.ModifiedBy || '' },
        { key: 'Modified', name: 'Modified', fieldName: 'Modified', minWidth: 100, maxWidth: 140, isResizable: true, onRender: (item) => formatDate(item.Modified) },
    ];
    return (
        <div className={styles.container}>
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    {items.length === 0 ? (
                        <div className={styles.emptyListMessage}>Nessun record disponibile</div>
                    ) : (
                        <DetailsList
                            items={items}
                            columns={columns}
                            layoutMode={DetailsListLayoutMode.justified}
                            constrainMode={ConstrainMode.unconstrained}
                        />
                    )}
                </div>
            )}
        </div>
    );
}