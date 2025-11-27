import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { DistributionListsService, UsersService } from '../../../services/SharePointService';
import { IDistributionListsProps, IDistributionListsItem } from './IDistributionListsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { UserHoverCardSmart } from '../UserHoverCard/UserHoverCard';

export const DistributionLists: React.FC<IDistributionListsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDistributionListsItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchDistributionLists = async (): Promise<void> => {
            try {
                const service = new DistributionListsService(context);
                const userService = new UsersService(context);
                const data = await service.getDistributionLists();
                const mapped: IDistributionListsItem[] = await Promise.all(data.map(async (item) => ({
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
                    CreatedBy: {
                        Id: item.Author ? item.Author.Id : 0,
                        Title: item.Author ? item.Author.Title : '',
                        EMail: item.Author?.EMail || undefined,
                        Picture: item.Author?.EMail ? await userService.getUserProfilePictureByEmail(item.Author.EMail) : undefined
                    },
                    ModifiedBy: {
                        Id: item.Editor ? item.Editor.Id : 0,
                        Title: item.Editor ? item.Editor.Title : '',
                        EMail: item.Editor?.EMail || undefined,
                        Picture: item.Editor?.EMail ? await userService.getUserProfilePictureByEmail(item.Editor.EMail) : undefined
                    }
                })));
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
        {
            key: 'CreatedBy',
            name: 'Created By',
            fieldName: 'CreatedBy',
            minWidth: 120,
            maxWidth: 200,
            isResizable: true,
            onRender: (item) => {
                if (!item.CreatedBy.EMail) return '';
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
            key: 'ModifiedBy',
            name: 'Modified By',
            fieldName: 'ModifiedBy',
            minWidth: 120,
            maxWidth: 200,
            isResizable: true,
            onRender: (item) => {
                if (!item.ModifiedBy.EMail) return '';
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