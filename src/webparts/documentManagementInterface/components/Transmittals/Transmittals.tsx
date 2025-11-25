import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { TransmittalsService } from '../../../services/SharePointService';
import { ITransmittalsProps, ITransmittalItem } from './ITransmittalsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';

export const Transmittals: React.FC<ITransmittalsProps> = ({ context }) => {
    const [items, setItems] = React.useState<ITransmittalItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchTransmittals = async (): Promise<void> => {
            try {
                const service = new TransmittalsService(context);
                const data = await service.getTransmittals();
                const stripHtml = (html: string) => html.replace(/<[^>]+>/g, '').trim();
                const mapped: ITransmittalItem[] = data.map((item) => ({
                    TransmittalNumber: item.TransmittalNumber || '',
                    ProjectCode: item.ProjectCode.ProjectCode || '',
                    ProjectTitle: item.ProjectCode.Title || '',
                    RecipientEmail: item.RecipientEmail || '',
                    RecipientName: item.RecipientName || '',
                    SenderEmail: item.SenderEmail || '',
                    SenderName: item.SenderName || '',
                    SentDate: item.SentDate || '',
                    Subject: item.Subject ? stripHtml(item.Notes) : '',
                    Status: item.Status || '',
                    Notes: item.Notes ? stripHtml(item.Notes) : '',
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author?.Title || '',
                    ModifiedBy: item.Editor?.Title || '',
                }));
                setItems(mapped);
                console.log('Fetched transmittals:', mapped);
            }
            catch {
                setItems([]);
            }
            finally {
                setLoading(false);
            }
        };
        void fetchTransmittals();
    }, []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };
    const columns: IColumn[] = [
        { key: 'TransmittalNumber', name: 'Transmittal Number', fieldName: 'TransmittalNumber', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ProjectCode', name: 'Project Code', fieldName: 'ProjectCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ProjectTitle', name: 'Project Title', fieldName: 'ProjectTitle', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'RecipienEmail', name: 'Recipient Email', fieldName: 'RecipientEmail', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'RecipientName', name: 'Recipient Name', fieldName: 'RecipientName', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'SenderEmail', name: 'Sender Email', fieldName: 'SenderEmail', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'SenderName', name: 'Sender Name', fieldName: 'SenderName', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'SentDate', name: 'Sent Date', fieldName: 'SentDate', onRender: (item) => formatDate(item.SentDate), minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'Subject', name: 'Subject', fieldName: 'Subject', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Notes', name: 'Notes', fieldName: 'Notes', minWidth: 120, maxWidth: 200, isResizable: true },
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
                    {items.length === 0 ? (
                        <div className={styles.emptyListMessage}>Nessun record disponibile</div>
                    ) : (
                        <DetailsList
                            items={items}
                            columns={columns}
                            setKey="multiple"
                            selectionMode={2}
                            selectionPreservedOnEmptyClick={true}
                            layoutMode={DetailsListLayoutMode.fixedColumns}
                            constrainMode={ConstrainMode.horizontalConstrained}
                            isHeaderVisible={true}
                            ariaLabelForSelectionColumn="Toggle selection"
                            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                            checkButtonAriaLabel="select row"
                        />
                    )}
                </div>
            )}
        </div>
    );
}
