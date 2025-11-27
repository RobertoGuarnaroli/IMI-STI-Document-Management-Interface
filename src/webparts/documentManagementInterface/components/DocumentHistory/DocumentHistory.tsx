import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { DocumentHistoryService, UsersService } from '../../../services/SharePointService';
import { IDocumentHistoryProps, IDocumentHistoryItem } from './IDocumentHistoryProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { UserHoverCardSmart } from '../UserHoverCard/UserHoverCard';

export const DocumentHistory: React.FC<IDocumentHistoryProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDocumentHistoryItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchDocumentHistory = async (): Promise<void> => {
            try {
                const service = new DocumentHistoryService(context);
                const userService = new UsersService(context);
                const data = await service.getDocumentHistory();
                const mapped: IDocumentHistoryItem[] = await Promise.all(data.map(async (item) => ({
                    DocumentId: item.DocumentId?.Id || '',
                    DocumentCode: item.DocumentId?.DocumentCode || '',
                    Revision: item.DocumentId?.Revision || '',
                    Action: item.Action || '',
                    PerformedBy: item.PerformedBy ? {
                        Id: item.PerformedBy.Id,
                        Title: item.PerformedBy.Title,
                        EMail: item.PerformedBy.EMail,
                        Picture: item.PerformedBy.EMail ? await userService.getUserProfilePictureByEmail(item.PerformedBy.EMail) : undefined
                    } : undefined,
                    ActionDate: item.ActionDate || '',
                    Status: item.Status || '',
                    ApprovalCode: item.ApprovalCode || '', 
                })));
                setItems(mapped);
                console.log('Fetched document history:', mapped);
            }
            catch {
                setItems([]);
            }
            finally {
                setLoading(false);
            }
        }
        void fetchDocumentHistory();
    }, []);
    const formatDate = (dateStr?: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };

    const columns: IColumn[] = [
        { key: 'DocumentID', name: 'Document ID', fieldName: 'DocumentId', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'DocumentTitle', name: 'Document Title', fieldName: 'DocumentCode', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Revision', name: 'Revision', fieldName: 'Revision', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Action', name: 'Action', fieldName: 'Action', minWidth: 100, maxWidth: 150, isResizable: true },
        {
            key: 'PerformedBy',
            name: 'Performed By',
            fieldName: 'PerformedBy',
            minWidth: 160,
            maxWidth: 220,
            isResizable: true,
            onRender: (item) => {
                if (!item.PerformedBy.EMail) return '';
                return (
                    <UserHoverCardSmart
                        email={item.PerformedBy.EMail}
                        displayName={item.PerformedBy.Title}
                        pictureUrl={item.PerformedBy.Picture}
                        context={context}
                    />
                );
            }
        },
        { key: 'ActionDate', name: 'Action Date', fieldName: 'ActionDate', minWidth: 100, maxWidth: 150, isResizable: true, onRender: item => formatDate(item.ActionDate) },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ApprovalCode', name: 'Approval Code', fieldName: 'ApprovalCode', minWidth: 100, maxWidth: 150, isResizable: true },
    ];

    if (loading) {
        return <LoadingSpinner />;
    }
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