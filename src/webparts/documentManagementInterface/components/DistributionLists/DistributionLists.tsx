import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { DistributionListsService } from '../../../services/SharePointService';
import { IDistributionListsProps, IDistributionListsItem } from './IDistributionListsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';

export const DistributionLists: React.FC<IDistributionListsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDistributionListsItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchDistributionLists = async (): Promise<void> => {
            try {
                const service = new DistributionListsService(context);
                const data = await service.getDistributionLists();
                const mapped: IDistributionListsItem[] = data.map((item) => ({
                    ProjectCode: item.ProjectCode.ProjectCode || '',
                    ContactEmail: item.ContactEmail || '',
                    Company: item.Company || '',
                    Role: item.Role || '',
                    ReceiveTransmittals: item.ReceiveTransmittals || false,
                    ReceiveNotifications: item.ReceiveNotifications || false,
                    ReceiveReminders: item.ReceiveReminders || false,
                    IsActive: item.IsActive || false,
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author ? { Id: item.Author.Id, Title: item.Author.Title } : undefined,
                    ModifiedBy: item.Editor ? { Id: item.Editor.Id, Title: item.Editor.Title } : undefined,
                }));
                setItems(mapped);
                console.log('Fetched distribution lists:', mapped);
            } catch (error) {
                setItems([]);
            } finally {
                setLoading(false);
            }
        };
        void fetchDistributionLists();
    }, []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };
    const columns: IColumn[] = [
        { key: 'ProjectCode', name: 'Project Code', fieldName: 'ProjectCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ContactEmail', name: 'Contact Email', fieldName: 'ContactEmail', minWidth: 150, maxWidth: 250, isResizable: true },
        { key: 'Company', name: 'Company', fieldName: 'Company', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Role', name: 'Role', fieldName: 'Role', minWidth: 100, maxWidth: 180, isResizable: true },
        { key: 'ReceiveTransmittals', name: 'Receive Transmittals', fieldName: 'ReceiveTransmittals', minWidth: 120, maxWidth: 160, isResizable: true, onRender: (item) => item.ReceiveTransmittals ? 'Yes' : 'No' },
        { key: 'ReceiveNotifications', name: 'Receive Notifications', fieldName: 'ReceiveNotifications', minWidth: 120, maxWidth: 160, isResizable: true, onRender: (item) => item.ReceiveNotifications ? 'Yes' : 'No' },
        { key: 'ReceiveReminders', name: 'Receive Reminders', fieldName: 'ReceiveReminders', minWidth: 120, maxWidth: 160, isResizable: true, onRender: (item) => item.ReceiveReminders ? 'Yes' : 'No' },
        { key: 'IsActive', name: 'Is Active', fieldName: 'IsActive', minWidth: 80, maxWidth: 120, isResizable: true, onRender: (item) => item.IsActive ? 'Yes' : 'No' },
        { key: 'CreatedBy', name: 'Created By', fieldName: 'CreatedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.CreatedBy ? item.CreatedBy.Title : '' },
        { key: 'ModifiedBy', name: 'Modified By', fieldName: 'ModifiedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.ModifiedBy ? item.ModifiedBy.Title : '' },
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