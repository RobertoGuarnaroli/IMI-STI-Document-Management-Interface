import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { AlertsService } from '../../../services/SharePointService';
import { IAlertsProps, IAlertItem } from './IAlertsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';

export const Alerts: React.FC<IAlertsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IAlertItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchAlerts = async (): Promise<void> => {
            try {
                const service = new AlertsService(context);
                const data = await service.getAlerts();
                const mapped: IAlertItem[] = data.map((item) => ({
                    Id: item.Id,
                    ProjectCode: item.ProjectCode ? {
                        Id: item.ProjectCode.Id,
                        ProjectCode: item.ProjectCode.ProjectCode || ''
                    } : undefined,
                    DocumentId: item.DocumentId ? {
                        Id: item.DocumentId.Id,
                        DocumentCode: item.DocumentId.DocumentCode || ''
                    } : undefined,
                    AlertType: item.AlertType || '',
                    Priority: item.Priority || '',
                    DaysOverdue: item.DaysOverdue || 0,
                    ExpectedDate: item.ExpectedDate || '',
                    Message: typeof item.Message === 'object' && item.Message !== null && 'value' in item.Message ? item.Message.value : (item.Message || ''),
                    AssignedTo: item.AssignedTo ? {
                        Id: item.AssignedTo.Id,
                        Title: item.AssignedTo.Title,
                        EMail: item.AssignedTo.EMail
                    } : undefined,
                    IsResolved: item.IsResolved || false,
                    ResolvedDate: item.ResolvedDate || '',
                }));
                setItems(mapped);
                console.log('Fetched alerts:', mapped);
            }
            catch {
                setItems([]);
            }
            finally {
                setLoading(false);
            }
        };
        void fetchAlerts();
    }, []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };

    // Utility to strip HTML tags from a string
    const stripHtml = (html: string) => {
        if (!html) return '';
        return html.replace(/<[^>]+>/g, '');
    };
    const columns: IColumn[] = [
        {
            key: 'ProjectCode',
            name: 'Project Code',
            fieldName: 'ProjectCode',
            minWidth: 80,
            maxWidth: 120,
            isResizable: true,
            onRender: (item) => item.ProjectCode ? item.ProjectCode.ProjectCode : ''
        },
        {
            key: 'DocumentId',
            name: 'Document ID',
            fieldName: 'DocumentId',
            minWidth: 80,
            maxWidth: 120,
            isResizable: true,
            onRender: (item) => item.DocumentId ? item.DocumentId.DocumentCode : ''
        },
        { key: 'AlertType', name: 'Alert Type', fieldName: 'AlertType', minWidth: 100, maxWidth: 150, isResizable: true },
        { key: 'Priority', name: 'Priority', fieldName: 'Priority', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'DaysOverdue', name: 'Days Overdue', fieldName: 'DaysOverdue', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ExpectedDate', name: 'Expected Date', fieldName: 'ExpectedDate', minWidth: 100, maxWidth: 150, isResizable: true, onRender: (item) => formatDate(item.ExpectedDate) },
        { key: 'Message', name: 'Message', fieldName: 'Message', minWidth: 200, maxWidth: 300, isResizable: true, onRender: (item) => stripHtml(item.Message) },
        { key: 'AssignedTo', name: 'Assigned To', fieldName: 'AssignedTo.Title', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.AssignedTo ? item.AssignedTo.Title : '' },
        { key: 'IsResolved', name: 'Resolved', fieldName: 'IsResolved', minWidth: 80, maxWidth: 100, isResizable: true, onRender: (item) => item.IsResolved ? 'Yes' : 'No' },
        { key: 'ResolvedDate', name: 'Resolved Date', fieldName: 'ResolvedDate', minWidth: 100, maxWidth: 150, isResizable: true, onRender: (item) => formatDate(item.ResolvedDate) },
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
