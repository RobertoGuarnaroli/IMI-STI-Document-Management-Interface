import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { DocumentHistoryService } from '../../../services/SharePointService';
import { IDocumentHistoryProps, IDocumentHistoryItem } from './IDocumentHistoryProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';

export const DocumentHistory: React.FC<IDocumentHistoryProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDocumentHistoryItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    React.useEffect(() => {
        const fetchDocumentHistory = async (): Promise<void> => {
            try {
                const service = new DocumentHistoryService(context);
                const data = await service.getDocumentHistory();
                const mapped: IDocumentHistoryItem[] = data.map((item) => ({
                    DocumentId: item.DocumentId?.Id || '',
                    DocumentCode: item.DocumentId?.DocumentCode || '',
                    Revision: item.DocumentId?.Revision || '',
                    Action: item.Action || '',
                    PerformedBy: item.PerformedBy ? {
                        Id: item.PerformedBy.Id,
                        Title: item.PerformedBy.Title,
                        Email: item.PerformedBy.EMail
                    } : undefined,
                    ActionDate: item.ActionDate || '',
                    Status: item.Status || '',
                    ApprovalCode: item.ApprovalCode || '', 
                }));
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
    const getPersonaPhotoUrl = (email?: string) => {
        if (!email) return undefined;
        return `https://outlook.office365.com/owa/service.svc/s/GetPersonaPhoto?email=${encodeURIComponent(email)}&size=HR96x96`;
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
            onRender: item =>
                item.PerformedBy ? (
                    <Persona
                        text={item.PerformedBy.Title}
                        imageUrl={getPersonaPhotoUrl(item.PerformedBy.EMail)}
                        size={PersonaSize.size32}
                        hidePersonaDetails={false}
                        secondaryText={item.PerformedBy.EMail}
                    />
                ) : ''
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
                            setKey="documentHistoryList"
                            layoutMode={DetailsListLayoutMode.fixedColumns}
                            constrainMode={ConstrainMode.unconstrained}
                            selectionMode={0}
                        />
                    )}
                </div>
            )}  
        </div>
    );
}