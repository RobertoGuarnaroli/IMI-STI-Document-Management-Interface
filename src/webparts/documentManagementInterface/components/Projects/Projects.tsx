import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { ProjectsService, UsersService } from '../../../services/SharePointService';
import { IProjectsProps, IProjectItem } from './IProjectsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
import { ExcelUpload, IExcelData } from '../ExcelUpload/ExcelUpload';
import { PeoplePicker } from '../PeoplePicker/PeoplePicker';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { ModalContainer } from '../ModalContainer/ModalContainer';
import { ErrorPopUp } from '../ErrorPopUp/ErrorPopUp';
import { ButtonsRibbon } from '../ButtonsRibbon/ButtonRibbons';
import { Selection } from '@fluentui/react/lib/Selection';

export const Projects: React.FC<IProjectsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IProjectItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [isModalOpen, setIsModalOpen] = React.useState(false);
    const [isEditModalOpen, setIsEditModalOpen] = React.useState(false);
    const [saving, setSaving] = React.useState(false);
    const [formError, setFormError] = React.useState<string | null>(null);
    const [dateError, setDateError] = React.useState<string | null>(null);
    const [showError, setShowError] = React.useState(false);
    const [touched, setTouched] = React.useState(false);
    const [showDeleteConfirm, setShowDeleteConfirm] = React.useState(false);
    const [statusOptions, setStatusOptions] = React.useState<{ key: string; text: string }[]>([]);
    const [selectedItems, setSelectedItems] = React.useState<IProjectItem[]>([]);
    const [editProjectId, setEditProjectId] = React.useState<number | null>(null);
    // Ref per la selezione della lista
    const selectionRef = React.useRef<Selection>(
        new Selection({
            onSelectionChanged: () => {
                const sel = selectionRef.current.getSelection() as IProjectItem[];
                setSelectedItems(sel);
            }
        })
    );
    const [newProject, setNewProject] = React.useState({
        ProjectCode: '',
        Title: '',
        Customer: '',
        ProjectManagerId: undefined as number | undefined,
        ProjectManagerTitle: '',
        Status: '',
        StartDate: '',
        EndDate: '',
        Notes: ''
    });
    const [editProject, setEditProject] = React.useState({
        Id: undefined as number | undefined,
        ProjectCode: '',
        Title: '',
        Customer: '',
        ProjectManagerId: undefined as number | undefined,
        ProjectManagerTitle: '',
        Status: '',
        StartDate: '',
        EndDate: '',
        Notes: ''
    });
    // Funzione fetchProjects a livello di componente
    const fetchProjects = async (): Promise<void> => {
        try {
            const service = new ProjectsService(context);
            const userProfileService = new (await import(/* webpackChunkName: 'SharePointService' */ '../../../services/SharePointService')).UserProfileService(context);
            const data = await service.getProjects();
            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
            const mapped: IProjectItem[] = await Promise.all(data.map(async (item) => {
                let projectManagerPicture = undefined;
                if (item.ProjectManager && item.ProjectManager.Id) {
                    projectManagerPicture = await userProfileService.getUserProfilePicture(item.ProjectManager.Id);
                }
                return {
                    key: item.Id,
                    Id: item.Id,
                    ProjectCode: item.ProjectCode || '',
                    Title: item.Title || '',
                    Customer: item.Customer || '',
                    ProjectManager: {
                        Title: item.ProjectManager?.Title || '',
                        Id: item.ProjectManager?.Id || undefined,
                        Picture: projectManagerPicture
                    },
                    Status: item.Status || '',
                    StartDate: item.StartDate || '',
                    EndDate: item.EndDate || '',
                    Notes: item.Notes ? stripHtml(item.Notes) : '',
                    Modified: item.Modified || '',
                    Created: item.Created || '',
                    CreatedBy: item.Author?.Title || '',
                    ModifiedBy: item.Editor?.Title || '',
                    context: context
                };
            }));
            setItems(mapped);
            console.log('Fetched projects:', mapped);
        } catch {
            setItems([]);
        } finally {
            setLoading(false);
        }
    };

    const handleSaveProject = async (): Promise<void> => {
        setTouched(true);
        setDateError(null);
        setFormError(null);
        setShowError(false);
        // Validazione campi obbligatori
        if (!newProject.ProjectCode || !newProject.Title || !newProject.Customer || !newProject.Status || !newProject.ProjectManagerId) {
            setFormError('Compila tutti i campi obbligatori prima di salvare.');
            setShowError(true);
            return;
        }
        if (newProject.StartDate && newProject.EndDate) {
            const start = new Date(newProject.StartDate);
            const end = new Date(newProject.EndDate);
            if (end < start) {
                setDateError('La data di fine non può essere precedente alla data di inizio.');
                setShowError(true);
                return;
            }
        }
        setSaving(true);
        try {
            const service = new ProjectsService(context);
            if (editProjectId) {
                await service.updateProject(editProjectId, {
                    ProjectCode: newProject.ProjectCode,
                    Title: newProject.Title,
                    Customer: newProject.Customer,
                    ProjectManagerId: newProject.ProjectManagerId,
                    Status: newProject.Status,
                    StartDate: newProject.StartDate,
                    EndDate: newProject.EndDate,
                    Notes: newProject.Notes
                });
            } else {
                await service.addProject({
                    ProjectCode: newProject.ProjectCode,
                    Title: newProject.Title,
                    Customer: newProject.Customer,
                    ProjectManagerId: newProject.ProjectManagerId,
                    Status: newProject.Status,
                    StartDate: newProject.StartDate,
                    EndDate: newProject.EndDate,
                    Notes: newProject.Notes
                });
            }
            setIsModalOpen(false);
            setEditProjectId(null);
            setNewProject({ ProjectCode: '', Title: '', Customer: '', ProjectManagerId: undefined, ProjectManagerTitle: '', Status: '', StartDate: '', EndDate: '', Notes: '' });
            setLoading(true);
            // Ricarica la lista
            await fetchProjects();
        } catch (error: unknown) {
            // Mostra solo il messaggio principale di errore
            const errorMsg = 'Errore durante la creazione del progetto.';
            if (typeof error === 'object' && error !== null) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const err = error as any;
                // Caso SharePoint REST (PnPjs): estrai solo il messaggio leggibile
                if (err.data?.odata?.error?.message?.value) {/* Lines 156-158 omitted */} else if (err.message && typeof err.message === 'string') {/* Lines 159-170 omitted */} else if (err.responseText && typeof err.responseText === 'string') {/* Lines 171-173 omitted */}
            }
            // Mostra solo la parte "leggibile" (senza prefissi tecnici)
            if (errorMsg.includes('::>')) {
                // Es: "Error making HttpClient request in queryable [500] ::> {json}"
                const match = errorMsg.match(/\{"odata.error":\{"code":"[^"]+","message":\{"lang":"[^"]+","value":"([^"]+)"/);
                if (match && match[1]) {/* Lines 180-181 omitted */}
            }
            setFormError(errorMsg);
            setShowError(true);
        } finally {
            setSaving(false);
            setLoading(false);
        }
    };


    const handleDeleteSelected = async (): Promise<void> => {
        setSaving(true);
        try {
            const service = new ProjectsService(context);
            for (const item of selectedItems) {
                if (item.Id) {
                    await service.deleteProject(item.Id);
                }
            }
            // Aggiorna la lista dopo l'eliminazione
            const data = await service.getProjects();
            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
            const mapped: IProjectItem[] = data.map((item) => ({
                key: item.Id,
                Id: item.Id,
                ProjectCode: item.ProjectCode || '',
                Title: item.Title || '',
                Customer: item.Customer || '',
                ProjectManager: {
                    Title: item.ProjectManager?.Title || '',
                    Id: item.ProjectManager?.Id || undefined,
                    Picture: item.ProjectManager?.Picture || undefined
                },
                Status: item.Status || '',
                StartDate: item.StartDate || '',
                EndDate: item.EndDate || '',
                Notes: item.Notes ? stripHtml(item.Notes) : '',
                Modified: item.Modified || '',
                Created: item.Created || '',
                CreatedBy: item.Author?.Title || '',
                ModifiedBy: item.Editor?.Title || '',
                context: context
            }));
            setItems(mapped);
            setSelectedItems([]);
            setShowDeleteConfirm(false);
        } catch {
            // TODO: gestione errore
        } finally {
            setSaving(false);
        }
    };

    React.useEffect(() => {
        void fetchProjects();
    }, []);

    React.useEffect(() => {
        const fetchStatusOptions = async (): Promise<void> => {
            const service = new ProjectsService(context);
            const choices = await service.getStatusChoices();
            setStatusOptions(choices.map(choice => ({ key: choice, text: choice })));
        };
        void fetchStatusOptions();
    }, [context]);

    // Imposta lo stato predefinito e unico per nuovo progetto
    React.useEffect(() => {
        if (isModalOpen && !editProjectId) {
            setNewProject(p => ({ ...p, Status: 'Active' }));
            setStatusOptions([{ key: 'Active', text: 'Active' }]);
        }
    }, [isModalOpen, editProjectId]);

    const formatDate = (dateStr?: string): string => {
        if (!dateStr) return '';
        const d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        return d.toLocaleDateString();
    };

    const columns: IColumn[] = [
        { key: 'ProjectCode', name: 'Project Code', fieldName: 'ProjectCode', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Customer', name: 'Customer', fieldName: 'Customer', minWidth: 80, maxWidth: 120, isResizable: true },
        {
            key: 'ProjectManager',
            name: 'Project Manager',
            fieldName: 'ProjectManager',
            minWidth: 80,
            maxWidth: 120,
            isResizable: true,
            onRender: (item) => (
                item.ProjectManager.Title ? (
                    <Persona
                        text={item.ProjectManager.Title}
                        size={PersonaSize.size32}
                        imageUrl={item.ProjectManager.Picture}
                    />
                ) : ''
            )
        },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'StartDate', name: 'Start Date', fieldName: 'StartDate', onRender: (item) => formatDate(item.StartDate), minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'EndDate', name: 'End Date', fieldName: 'EndDate', onRender: (item) => formatDate(item.EndDate), minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'Notes', name: 'Notes', fieldName: 'Notes', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'CreatedBy', name: 'Created By', fieldName: 'CreatedBy', minWidth: 80, maxWidth: 120, isResizable: true, onRender: (item) => item.CreatedBy || '' },
        { key: 'Created', name: 'Created', fieldName: 'Created', onRender: (item) => formatDate(item.Created), minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'ModifiedBy', name: 'Modified By', fieldName: 'ModifiedBy', minWidth: 80, maxWidth: 120, isResizable: true, onRender: (item) => item.ModifiedBy || '' },
        { key: 'Modified', name: 'Modified', fieldName: 'Modified', onRender: (item) => formatDate(item.Modified), minWidth: 80, maxWidth: 120, isResizable: true },
    ];

    return (
        <div className={styles.container}>
            <ButtonsRibbon
                buttons={
                    [
                        {
                            key: 'addProject',
                            text: 'New Project',
                            iconName: 'Add',
                            onClick: () => {
                                setNewProject({
                                    ProjectCode: '',
                                    Title: '',
                                    Customer: '',
                                    ProjectManagerId: undefined,
                                    ProjectManagerTitle: '',
                                    Status: '',
                                    StartDate: '',
                                    EndDate: '',
                                    Notes: ''
                                });
                                setTouched(false);
                                setIsModalOpen(true);
                            },
                            disabled: false,
                            color: '#5a2a6b',
                            border: '#5a2a6b'
                        },
                        selectedItems.length === 1 ? {
                            key: 'editProject',
                            text: 'Edit Project',
                            iconName: 'Edit',
                            onClick: () => {
                                const item = selectedItems[0];
                                setEditProject({
                                    Id: item.Id,
                                    ProjectCode: item.ProjectCode || '',
                                    Title: item.Title || '',
                                    Customer: item.Customer || '',
                                    ProjectManagerId: item.ProjectManager.Id ? Number(item.ProjectManager.Id) : undefined,
                                    ProjectManagerTitle: item.ProjectManager.Title || '',
                                    Status: item.Status || '',
                                    StartDate: item.StartDate || '',
                                    EndDate: item.EndDate || '',
                                    Notes: item.Notes || ''
                                });
                                setIsEditModalOpen(true);
                            },
                            disabled: false,
                            color: '#5a2a6b',
                            border: '#5a2a6b'
                        } : null,
                        selectedItems.length > 0 ? {
                            key: 'deleteProject',
                            text: 'Delete Project',
                            iconName: 'Delete',
                            onClick: () => setShowDeleteConfirm(true),
                            disabled: false,
                            color: '#a4262c',
                            border: '#a4262c'
                        } : null
                    ].filter((b): b is NonNullable<typeof b> => b !== null)
                }
            />
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
                            selection={selectionRef.current}
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
            {isModalOpen && (
                <ModalContainer
                    isOpen={isModalOpen}
                    title="Crea nuovo Project"
                    onSave={handleSaveProject}
                    onCancel={() => {
                        setNewProject({
                            ProjectCode: '',
                            Title: '',
                            Customer: '',
                            ProjectManagerId: undefined,
                            ProjectManagerTitle: '',
                            Status: '',
                            StartDate: '',
                            EndDate: '',
                            Notes: ''
                        });
                        setTouched(false);
                        setIsModalOpen(false);
                    }}
                    saving={saving}
                >
                    <ExcelUpload
                        onDataExtracted={async (data: IExcelData) => {
                            setNewProject(p => ({
                                ...p,
                                ProjectCode: data.ProjectCode || '',
                                Title: data.Title || '',
                                Customer: data.Customer || '',
                                Status: data.Status || '',
                                StartDate: data.StartDate || '',
                                EndDate: data.EndDate || '',
                                Notes: data.Notes || ''
                            }));
                            if (data.ProjectManager) {
                                const usersService = new UsersService(context);
                                const users = await usersService.getUsers();
                                const found = users.find(u => u.secondaryText && u.secondaryText.toLowerCase() === data.ProjectManager!.toLowerCase());
                                if (found) {
                                    setNewProject(p => ({
                                        ...p,
                                        ProjectManagerId: found.id ? Number(found.id) : undefined,
                                        ProjectManagerTitle: found.text || ''
                                    }));
                                }
                            }
                        }}
                        onClear={() => {
                            setNewProject({
                                ProjectCode: '',
                                Title: '',
                                Customer: '',
                                ProjectManagerId: undefined,
                                ProjectManagerTitle: '',
                                Status: '',
                                StartDate: '',
                                EndDate: '',
                                Notes: ''
                            });
                        }}
                        disabled={saving}
                    />
                    <TextField label="Project Code" value={newProject.ProjectCode} onChange={(_, v) => setNewProject(p => ({ ...p, ProjectCode: v || '' }))} required />
                    <TextField label="Title" value={newProject.Title} onChange={(_, v) => setNewProject(p => ({ ...p, Title: v || '' }))} required />
                    <TextField label="Customer" value={newProject.Customer} onChange={(_, v) => setNewProject(p => ({ ...p, Customer: v || '' }))} required />
                    <Dropdown
                        label="Status"
                        options={statusOptions}
                        selectedKey={newProject.Status}
                        onChange={(_, option) => setNewProject(p => ({ ...p, Status: option?.key as string }))}
                        required
                        placeholder="Seleziona uno stato"
                    />
                    <PeoplePicker
                        key="new"
                        context={context}
                        titleText="Project Manager"
                        personSelectionLimit={1}
                        showtooltip={true}
                        required={true}
                        selectedUserIds={typeof newProject.ProjectManagerId === 'number' ? [newProject.ProjectManagerId] : []}
                        onChange={userIds => {
                            setNewProject(p => ({
                                ...p,
                                ProjectManagerId: userIds[0] ? Number(userIds[0]) : undefined
                            }));
                        }}
                        principalTypes={[1]}
                        resolveDelay={300}
                        ensureUser={true}
                        showHiddenInUI={false}
                        suggestionsLimit={15}
                        disabled={false}
                        label="Project Manager"
                        placeholder="Seleziona il Project Manager"
                        itemLimit={1}
                        loadUsers={async (context) => {
                            const service = new UsersService(context);
                            return await service.getUsers();
                        }}
                    />
                    <DatePicker
                        label="Start Date"
                        value={newProject.StartDate ? new Date(newProject.StartDate) : undefined}
                        onSelectDate={date => setNewProject(p => ({ ...p, StartDate: date ? date.toISOString().substring(0, 10) : '' }))}
                        placeholder="DD-MM-YYYY"
                        formatDate={d => d ? d.toLocaleDateString() : ''}
                        isRequired={touched}
                    />
                    <DatePicker
                        label="End Date"
                        value={newProject.EndDate ? new Date(newProject.EndDate) : undefined}
                        onSelectDate={date => setNewProject(p => ({ ...p, EndDate: date ? date.toISOString().substring(0, 10) : '' }))}
                        placeholder="DD-MM-YYYY"
                        formatDate={d => d ? d.toLocaleDateString() : ''}
                        isRequired={touched}
                    />
                    <TextField label="Notes" multiline rows={3} value={newProject.Notes} onChange={(_, v) => setNewProject(p => ({ ...p, Notes: v || '' }))} />
                </ModalContainer>
            )}

            {isEditModalOpen && (
                <ModalContainer
                    isOpen={isEditModalOpen}
                    title="Modifica Project"
                    onSave={async () => {
                        setDateError(null);
                        setFormError(null);
                        setShowError(false);
                        if (!editProject.ProjectCode || !editProject.Title || !editProject.Customer || !editProject.Status || !editProject.ProjectManagerId) {
                            setFormError('Compila tutti i campi obbligatori prima di salvare.');
                            setShowError(true);
                            return;
                        }
                        if (editProject.StartDate && editProject.EndDate) {
                            const start = new Date(editProject.StartDate);
                            const end = new Date(editProject.EndDate);
                            if (end < start) {
                                setDateError('La data di fine non può essere precedente alla data di inizio.');
                                setShowError(true);
                                return;
                            }
                        }
                        setSaving(true);
                        try {
                            const service = new ProjectsService(context);
                            await service.updateProject(editProject.Id!, {
                                ProjectCode: editProject.ProjectCode,
                                Title: editProject.Title,
                                Customer: editProject.Customer,
                                ProjectManagerId: editProject.ProjectManagerId,
                                Status: editProject.Status,
                                StartDate: editProject.StartDate,
                                EndDate: editProject.EndDate,
                                Notes: editProject.Notes
                            });
                            setIsEditModalOpen(false);
                            setEditProject({
                                Id: undefined,
                                ProjectCode: '',
                                Title: '',
                                Customer: '',
                                ProjectManagerId: undefined,
                                ProjectManagerTitle: '',
                                Status: '',
                                StartDate: '',
                                EndDate: '',
                                Notes: ''
                            });
                            setLoading(true);
                            const data = await service.getProjects();
                            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
                            const mapped: IProjectItem[] = data.map((item) => ({
                                key: item.Id,
                                Id: item.Id,
                                ProjectCode: item.ProjectCode || '',
                                Title: item.Title || '',
                                Customer: item.Customer || '',
                                ProjectManager: {
                                    Id: item.ProjectManager?.Id || undefined,
                                    Title: item.ProjectManager?.Title || '',
                                    Picture: item.ProjectManager?.Picture || undefined
                                },
                                Status: item.Status || '',
                                StartDate: item.StartDate || '',
                                EndDate: item.EndDate || '',
                                Notes: item.Notes ? stripHtml(item.Notes) : '',
                                Modified: item.Modified || '',
                                Created: item.Created || '',
                                CreatedBy: item.Author?.Title || '',
                                ModifiedBy: item.Editor?.Title || '',
                                context: context
                            }));
                            setItems(mapped);
                        } catch (error: unknown) {
                            let errorMsg = 'Errore durante la modifica del progetto.';
                            if (typeof error === 'object' && error !== null) {
                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                const err = error as any;
                                if (err.data?.odata?.error?.message?.value) {
                                    errorMsg = err.data.odata.error.message.value;
                                } else if (err.message && typeof err.message === 'string') {
                                    try {
                                        const parsed = JSON.parse(err.message);
                                        if (parsed?.odata?.error?.message?.value) {
                                            errorMsg = parsed.odata.error.message.value;
                                        } else {
                                            errorMsg = err.message;
                                        }
                                    } catch {
                                        errorMsg = err.message;
                                    }
                                } else if (err.responseText && typeof err.responseText === 'string') {
                                    errorMsg = err.responseText;
                                }
                            }
                            if (errorMsg.includes('::>')) {
                                const match = errorMsg.match(/\{"odata.error":\{"code":"[^"]+","message":\{"lang":"[^"]+","value":"([^"]+)"/);
                                if (match && match[1]) {
                                    errorMsg = match[1];
                                }
                            }
                            setFormError(errorMsg);
                            setShowError(true);
                        } finally {
                            setSaving(false);
                            setLoading(false);
                        }
                    }}
                    onCancel={() => {
                        setIsEditModalOpen(false);
                        setEditProject({
                            Id: undefined,
                            ProjectCode: '',
                            Title: '',
                            Customer: '',
                            ProjectManagerId: undefined,
                            ProjectManagerTitle: '',
                            Status: '',
                            StartDate: '',
                            EndDate: '',
                            Notes: ''
                        });
                    }}
                    saving={saving}
                >
                    <TextField label="Project Code" value={editProject.ProjectCode} onChange={(_, v) => setEditProject(p => ({ ...p, ProjectCode: v || '' }))} required />
                    <TextField label="Title" value={editProject.Title} onChange={(_, v) => setEditProject(p => ({ ...p, Title: v || '' }))} required />
                    <TextField label="Customer" value={editProject.Customer} onChange={(_, v) => setEditProject(p => ({ ...p, Customer: v || '' }))} required />
                    <Dropdown
                        label="Status"
                        options={statusOptions}
                        selectedKey={editProject.Status}
                        onChange={(_, option) => setEditProject(p => ({ ...p, Status: option?.key as string }))}
                        required
                        placeholder="Seleziona uno stato"
                    />
                    <PeoplePicker
                        key={editProject.Id ? `edit-${editProject.Id}` : 'edit'}
                        context={context}
                        titleText="Project Manager"
                        personSelectionLimit={1}
                        showtooltip={true}
                        required={true}
                        selectedUserIds={typeof editProject.ProjectManagerId === 'number' ? [editProject.ProjectManagerId] : []}
                        onChange={userIds => {
                            setEditProject(p => ({
                                ...p,
                                ProjectManagerId: userIds[0] ? Number(userIds[0]) : undefined
                            }));
                        }}
                        principalTypes={[1]}
                        resolveDelay={300}
                        ensureUser={true}
                        showHiddenInUI={false}
                        suggestionsLimit={15}
                        disabled={false}
                        label="Project Manager"
                        placeholder="Seleziona il Project Manager"
                        itemLimit={1}
                        loadUsers={async (context) => {
                            const service = new UsersService(context);
                            return await service.getUsers();
                        }}
                    />
                    <DatePicker
                        label="Start Date"
                        value={editProject.StartDate ? new Date(editProject.StartDate) : undefined}
                        onSelectDate={date => setEditProject(p => ({ ...p, StartDate: date ? date.toISOString().substring(0, 10) : '' }))}
                        placeholder="DD-MM-YYYY"
                        formatDate={d => d ? d.toLocaleDateString() : ''}
                        isRequired={true}
                    />
                    <DatePicker
                        label="End Date"
                        value={editProject.EndDate ? new Date(editProject.EndDate) : undefined}
                        onSelectDate={date => setEditProject(p => ({ ...p, EndDate: date ? date.toISOString().substring(0, 10) : '' }))}
                        placeholder="DD-MM-YYYY"
                        formatDate={d => d ? d.toLocaleDateString() : ''}
                        isRequired={true}
                    />
                    <TextField label="Notes" multiline rows={3} value={editProject.Notes} onChange={(_, v) => setEditProject(p => ({ ...p, Notes: v || '' }))} />
                </ModalContainer>
            )}
            {showError && (
                <ErrorPopUp
                    message={formError || dateError || ''}
                    onClose={() => setShowError(false)}
                    duration={4000}
                />
            )}
            {showDeleteConfirm && (
                <ModalContainer
                    isOpen={showDeleteConfirm}
                    title={selectedItems.length === 1 ? 'Conferma eliminazione progetto' : `Conferma eliminazione di ${selectedItems.length} progetti`}
                    onSave={handleDeleteSelected}
                    onCancel={() => setShowDeleteConfirm(false)}
                    saving={saving}
                    saveText="Elimina"
                    cancelText="Annulla"
                >
                    <div className={styles.deleteConfirmText}>
                        Sei sicuro di voler eliminare {selectedItems.length === 1 ? 'questo progetto' : `i ${selectedItems.length} progetti selezionati`}?
                    </div>
                </ModalContainer>
            )}
        </div>
    );
}