import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { TextField, Dropdown, IDropdownOption } from '@fluentui/react';
import { PeoplePicker } from '../PeoplePicker/PeoplePicker';
import { ButtonsRibbon } from '../ButtonsRibbon/ButtonRibbons';
import { ModalContainer } from '../ModalContainer/ModalContainer';
import { DocumentsService, UsersService, ChoiceFieldService } from '../../../services/SharePointService';
import { IDocumentsProps, IDocumentItem } from './IDocumentsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { UserHoverCardSmart } from '../UserHoverCard/UserHoverCard';

export const Documents: React.FC<IDocumentsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IDocumentItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [showModal, setShowModal] = React.useState(false);
    const [saving, setSaving] = React.useState(false);
    const [form, setForm] = React.useState({
        DocumentCode: '',
        Title: '',
        Revision: '',
        Status: '',
        IssuePurpose: '',
        ApprovalCode: '',
        SentDate: undefined as Date | undefined,
        ExpectedReturnDate: undefined as Date | undefined,
        ActualReturnDate: undefined as Date | undefined,
        AssignedToId: undefined as number | undefined,
        Notes: '',
        TurnaroundDays: 0,
        DaysLate: 0,
    });
    const [editMode, setEditMode] = React.useState(false);
    const [editItemId, setEditItemId] = React.useState<number | null>(null);
    const [formError, setFormError] = React.useState<string | null>(null);
    const [statusOptions, setStatusOptions] = React.useState<IDropdownOption[]>([]);
    const [issuePurposeOptions, setIssuePurposeOptions] = React.useState<IDropdownOption[]>([]);
    const [approvalCodeOptions, setApprovalCodeOptions] = React.useState<IDropdownOption[]>([]);
    React.useEffect(() => {
        const fetchChoices = async () => {
            const service = new ChoiceFieldService(context);
            try {
                const status = await service.getFieldChoices("Documents", "Status");
                setStatusOptions(status.map((c: string) => ({ key: c, text: c })));
                const issuePurpose = await service.getFieldChoices("Documents", "IssuePurpose");
                setIssuePurposeOptions(issuePurpose.map((c: string) => ({ key: c, text: c })));
                const approvalCode = await service.getFieldChoices("Documents", "ApprovalCode");
                setApprovalCodeOptions(approvalCode.map((c: string) => ({ key: c, text: c })));
            } catch {}
        };
        void fetchChoices();
    }, [context]);
    const stripHtml = (html: string) => html.replace(/<[^>]+>/g, '').trim();
    const extractDate = (dateStr: string) => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return '';
        return d.toISOString().slice(0, 10);
    };
    const fetchDocuments = async (): Promise<void> => {
        try {
            const service = new DocumentsService(context);
            const usersService = new UsersService(context);
            const data = await service.getDocuments();
            // Recupera le immagini profilo in parallelo
            const mapped: IDocumentItem[] = await Promise.all(data.map(async (item) => {
                return {
                    DocumentCode: item.DocumentCode || '',
                    Title: item.Title || '',
                    Revision: item.Revision || '',
                    Status: item.Status || '',
                    IssuePurpose: item.IssuePurpose || '',
                    ApprovalCode: item.ApprovalCode || '',
                    SentDate: extractDate(item.SentDate),
                    ExpectedReturnDate: extractDate(item.ExpectedReturnDate),
                    ActualReturnDate: extractDate(item.ActualReturnDate),
                    TurnaroundDays: Math.round(item.TurnaroundDays || 0),
                    DaysLate: Math.round(item.DaysLate || 0),
                    AssignedTo: {
                        Title: item.AssignedTo?.Title || '',
                        Id: item.AssignedTo?.Id || undefined,
                        EMail: item.AssignedTo?.EMail || '',
                        Picture: item.AssignedTo?.EMail ? await usersService.getUserProfilePictureByEmail(item.AssignedTo.EMail) : undefined
                    },
                    Notes: item.Notes ? stripHtml(item.Notes) : '',
                    Modified: extractDate(item.Modified),
                    Created: extractDate(item.Created),
                    CreatedBy: {
                        Title: item.Author?.Title || '',
                        Id: item.Author?.Id || undefined,
                        EMail: item.Author?.EMail || '',
                        Picture: item.Author?.EMail ? await usersService.getUserProfilePictureByEmail(item.Author.EMail) : undefined
                    },
                    ModifiedBy: {
                        Title: item.Editor?.Title || '',
                        Id: item.Editor?.Id || undefined,
                        EMail: item.Editor?.EMail || '',
                        Picture: item.Editor?.EMail ? await usersService.getUserProfilePictureByEmail(item.Editor.EMail) : undefined
                    },
                };
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

    React.useEffect(() => {
        void fetchDocuments();
    }, [context]); // End of useEffect
    // Define columns for DetailsList
    const columns: IColumn[] = [
        { key: 'DocumentCode', name: 'Document Code', fieldName: 'DocumentCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Revision', name: 'Revision', fieldName: 'Revision', minWidth: 60, maxWidth: 80, isResizable: true },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'IssuePurpose', name: 'Issue Purpose', fieldName: 'IssuePurpose', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ApprovalCode', name: 'Approval Code', fieldName: 'ApprovalCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'SentDate', name: 'Sent Date', fieldName: 'SentDate', minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'ExpectedReturnDate', name: 'Expected Return Date', fieldName: 'ExpectedReturnDate', minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'ActualReturnDate', name: 'Actual Return Date', fieldName: 'ActualReturnDate', minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'TurnaroundDays', name: 'Turnaround Days', fieldName: 'TurnaroundDays', minWidth: 60, maxWidth: 80, isResizable: true },
        { key: 'DaysLate', name: 'Days Late', fieldName: 'DaysLate', minWidth: 60, maxWidth: 80, isResizable: true },
        {
            key: 'AssignedTo',
            name: 'Assigned To',
            fieldName: 'AssignedTo',
            minWidth: 100,
            maxWidth: 160,
            isResizable: true,
            onRender: (item) => {
                if (!item.AssignedTo.EMail) return '';
                return (
                    <UserHoverCardSmart
                        email={item.AssignedTo.EMail}
                        displayName={item.AssignedTo.Title}
                        pictureUrl={item.AssignedTo.Picture}
                        context={context}
                    />
                );
            }
        },
        { key: 'Notes', name: 'Notes', fieldName: 'Notes', minWidth: 120, maxWidth: 200, isResizable: true },
        {
            key: 'CreatedBy',
            name: 'Created By',
            fieldName: 'CreatedBy',
            minWidth: 80,
            maxWidth: 120,
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
        { key: 'Created', name: 'Created', fieldName: 'Created', minWidth: 100, maxWidth: 140, isResizable: true },
        {
            key: 'ModifiedBy',
            name: 'Modified By',
            fieldName: 'ModifiedBy',
            minWidth: 80,
            maxWidth: 120,
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
        { key: 'Modified', name: 'Modified', fieldName: 'Modified', minWidth: 100, maxWidth: 140, isResizable: true },
    ];

    // Save document handler
    const handleSaveDocument = async () => {
        setFormError(null);
        if (!form.DocumentCode || !form.Title || !form.Revision) {
            setFormError('Compila tutti i campi obbligatori.');
            return;
        }
        setSaving(true);
        try {
            const service = new DocumentsService(context);
            if (editMode && editItemId !== null) {
                // Edit mode: update
                await service.updateDocument(editItemId, {
                    DocumentCode: form.DocumentCode,
                    Title: form.Title,
                    Revision: form.Revision,
                    Status: form.Status,
                    IssuePurpose: form.IssuePurpose,
                    ApprovalCode: form.ApprovalCode,
                    SentDate: form.SentDate ? form.SentDate.toISOString() : undefined,
                    ExpectedReturnDate: form.ExpectedReturnDate ? form.ExpectedReturnDate.toISOString() : undefined,
                    ActualReturnDate: form.ActualReturnDate ? form.ActualReturnDate.toISOString() : undefined,
                    TurnaroundDays: form.TurnaroundDays,
                    DaysLate: form.DaysLate,
                    AssignedToId: form.AssignedToId,
                    Notes: form.Notes,
                });
            } else {
                // New document
                await service.createDocument({
                    DocumentCode: form.DocumentCode,
                    Title: form.Title,
                    Revision: form.Revision,
                    Status: form.Status,
                    IssuePurpose: form.IssuePurpose,
                    ApprovalCode: form.ApprovalCode,
                    SentDate: form.SentDate ? form.SentDate.toISOString() : undefined,
                    ExpectedReturnDate: form.ExpectedReturnDate ? form.ExpectedReturnDate.toISOString() : undefined,
                    ActualReturnDate: form.ActualReturnDate ? form.ActualReturnDate.toISOString() : undefined,
                    TurnaroundDays: form.TurnaroundDays,
                    DaysLate: form.DaysLate,
                    AssignedToId: form.AssignedToId,
                    Notes: form.Notes,
                });
            }
            setShowModal(false);
            setForm({
                DocumentCode: '', Title: '', Revision: '', Status: '', IssuePurpose: '', ApprovalCode: '',
                SentDate: undefined, ExpectedReturnDate: undefined, ActualReturnDate: undefined,
                TurnaroundDays: 0, DaysLate: 0, AssignedToId: undefined, Notes: ''
            });
            setEditMode(false);
            setEditItemId(null);
            setSaving(false);
            setFormError(null);
            setLoading(true);
            // Refresh list
            await fetchDocuments();
        } catch (err) {
            setFormError('Errore durante il salvataggio.');
        } finally {
            setSaving(false);
        }
    };

    const loadUsers = async (): Promise<any[]> => {
        const usersService = new UsersService(context);
        const users = await usersService.getUsers();
        return users.map((u: any) => ({
            id: u.Id,
            text: u.Title,
            secondaryText: u.Email,
            imageUrl: u.Picture
        }));
    };

    return (
        <div className={styles.container}>
            <ButtonsRibbon
                buttons={[
                    {
                        key: 'newDocument',
                        text: 'New Document',
                        iconName: 'Add',
                        onClick: () => {
                            setForm({
                                DocumentCode: '', Title: '', Revision: '', Status: '', IssuePurpose: '', ApprovalCode: '',
                                SentDate: undefined, ExpectedReturnDate: undefined, ActualReturnDate: undefined,
                                TurnaroundDays: 0, DaysLate: 0, AssignedToId: undefined, Notes: ''
                            });
                            setEditMode(false);
                            setEditItemId(null);
                            setShowModal(true);
                        },
                        disabled: false,
                        color: '#5a2a6b',
                        border: '#5a2a6b'
                    },
                ]}
            />
            {showModal && (
                <ModalContainer
                    isOpen={showModal}
                    title="New Document"
                    onSave={handleSaveDocument}
                    onCancel={() => setShowModal(false)}
                    saving={saving}
                    width="600px"
                >
                    <TextField label="Document Code" required value={form.DocumentCode} onChange={(_, v) => setForm(f => ({ ...f, DocumentCode: v || '' }))} />
                    <TextField label="Title" required value={form.Title} onChange={(_, v) => setForm(f => ({ ...f, Title: v || '' }))} />
                    <TextField label="Revision" required value={form.Revision} onChange={(_, v) => setForm(f => ({ ...f, Revision: v || '' }))} />
                    <Dropdown
                        label="Status"
                        options={statusOptions}
                        selectedKey={form.Status}
                        onChange={(_, option) => setForm(f => ({ ...f, Status: option?.key as string }))}
                        required
                        placeholder="Seleziona uno stato"
                    />
                    <Dropdown
                        label="Issue Purpose"
                        options={issuePurposeOptions}
                        selectedKey={form.IssuePurpose}
                        onChange={(_, option) => setForm(f => ({ ...f, IssuePurpose: option?.key as string }))}
                        required
                        placeholder="Seleziona uno scopo"
                    />
                    <Dropdown
                        label="Approval Code"
                        options={approvalCodeOptions}
                        selectedKey={form.ApprovalCode}
                        onChange={(_, option) => setForm(f => ({ ...f, ApprovalCode: option?.key as string }))}
                        required
                        placeholder="Seleziona un codice"
                    />
                    
                    <PeoplePicker
                        key="assignedTo"
                        context={context}
                        titleText="Assigned To"
                        personSelectionLimit={1}
                        showtooltip={true}
                        required={false}
                        selectedUserIds={typeof form.AssignedToId === 'number' ? [form.AssignedToId] : []}
                        onChange={userIds => setForm(f => ({ ...f, AssignedToId: userIds[0] ? Number(userIds[0]) : undefined }))}
                        principalTypes={[1]}
                        resolveDelay={300}
                        ensureUser={true}
                        showHiddenInUI={false}
                        suggestionsLimit={15}
                        disabled={false}
                        label="Assigned To"
                        placeholder="Select user"
                        itemLimit={1}
                        loadUsers={loadUsers}
                    />
                    <TextField label="Notes" multiline value={form.Notes} onChange={(_, v) => setForm(f => ({ ...f, Notes: v || '' }))} />
                    {formError && <div style={{ color: 'red', marginTop: 8 }}>{formError}</div>}
                </ModalContainer>
            )}
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    {items.length === 0 ? (
                        <div className={styles.emptyListMessage}>Nessun record disponibile</div>
                    ) : (
                        <DetailsList
                            items={items}
                            columns={[
                                ...columns,
                                {
                                    key: 'edit',
                                    name: '',
                                    fieldName: 'edit',
                                    minWidth: 32,
                                    maxWidth: 40,
                                    isResizable: false,
                                    onRender: (item, idx) => (
                                        <span style={{ cursor: 'pointer', color: '#5a2a6b' }} title="Edit"
                                            onClick={() => {
                                                setForm({
                                                    DocumentCode: item.DocumentCode,
                                                    Title: item.Title,
                                                    Revision: item.Revision,
                                                    Status: item.Status,
                                                    IssuePurpose: item.IssuePurpose,
                                                    ApprovalCode: item.ApprovalCode,
                                                    SentDate: item.SentDate ? new Date(item.SentDate) : undefined,
                                                    ExpectedReturnDate: item.ExpectedReturnDate ? new Date(item.ExpectedReturnDate) : undefined,
                                                    ActualReturnDate: item.ActualReturnDate ? new Date(item.ActualReturnDate) : undefined,
                                                    TurnaroundDays: item.TurnaroundDays,
                                                    DaysLate: item.DaysLate,
                                                    AssignedToId: item.AssignedTo?.Id,
                                                    Notes: item.Notes,
                                                });
                                                setEditMode(true);
                                                setEditItemId(idx !== undefined ? idx + 1 : null);
                                                setShowModal(true);
                                            }}>
                                            ✏️
                                        </span>
                                    )
                                }
                            ]}
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
};