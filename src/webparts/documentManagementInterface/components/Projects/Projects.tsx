import * as React from 'react';
import { DetailsList, IColumn, DetailsListLayoutMode, ConstrainMode } from '@fluentui/react/lib/DetailsList';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { ProjectsService, UsersService } from '../../../services/SharePointService';
import { IProjectsProps, IProjectItem } from './IProjectsProps';
import styles from '../../styles/TabStyle.module.scss';
import { LoadingSpinner } from '../Spinner/Spinner';
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

    // Stato per la modifica
    const [editProjectId, setEditProjectId] = React.useState<number | null>(null);
    const [statusOptions, setStatusOptions] = React.useState<IDropdownOption[]>([]);
    const [saving, setSaving] = React.useState(false);
    const [dateError, setDateError] = React.useState<string | null>(null);
    const [showError, setShowError] = React.useState(false);
    const [formError, setFormError] = React.useState<string | null>(null);
    const selectionRef = React.useRef<Selection | null>(null);
    const [selectedItems, setSelectedItems] = React.useState<IProjectItem[]>([]);
    const [showDeleteConfirm, setShowDeleteConfirm] = React.useState(false);

    React.useEffect(() => {
        if (!selectionRef.current) {
            selectionRef.current = new Selection({
                onSelectionChanged: () => {
                    const selected = selectionRef.current?.getSelection() as IProjectItem[];
                    setSelectedItems(selected);
                }
            });
        }
    }, []);

    React.useEffect(() => {
        if (selectionRef.current) {
            selectionRef.current.setItems(items, true);
        }
    }, [items]);

    const handleSaveProject = async () => {
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
            const data = await service.getProjects();
            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
            const mapped: IProjectItem[] = data.map((item) => ({
                key: item.Id,
                Id: item.Id,
                ProjectCode: item.ProjectCode || '',
                Title: item.Title || '',
                Customer: item.Customer || '',
                ProjectManagerTitle: item.ProjectManager?.Title || '',
                ProjectManagerId: item.ProjectManager?.Id || undefined,
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
            // Mostra solo il messaggio principale di errore
            let errorMsg = 'Errore durante la creazione del progetto.';
            if (typeof error === 'object' && error !== null) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const err = error as any;
                // Caso SharePoint REST (PnPjs): estrai solo il messaggio leggibile
                if (err.data?.odata?.error?.message?.value) {
                    // Esempio: "The list item could not be added or updated because duplicate values were found in the following field(s) in the list: [ProjectCode]."
                    errorMsg = err.data.odata.error.message.value;
                } else if (err.message && typeof err.message === 'string') {
                    // Se il messaggio contiene JSON, estrai solo la parte leggibile
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
                    // Alcuni errori custom
                    errorMsg = err.responseText;
                }
            }
            // Mostra solo la parte "leggibile" (senza prefissi tecnici)
            if (errorMsg.includes('::>')) {
                // Es: "Error making HttpClient request in queryable [500] ::> {json}"
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
                ProjectManagerTitle: item.ProjectManager?.Title || '',
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
        const fetchProjects = async (): Promise<void> => {
            try {
                const service = new ProjectsService(context);
                const data = await service.getProjects();
                const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
                const mapped: IProjectItem[] = data.map((item) => ({
                    key: item.Id,
                    Id: item.Id,
                    ProjectCode: item.ProjectCode || '',
                    Title: item.Title || '',
                    Customer: item.Customer || '',
                    ProjectManagerTitle: item.ProjectManager?.Title || '',
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
                console.log('Fetched projects:', mapped);
            } catch {
                setItems([]);
            } finally {
                setLoading(false);
            }
        };
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
        { key: 'Title', name: 'Title', fieldName: 'Title', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'Customer', name: 'Customer', fieldName: 'Customer', minWidth: 120, maxWidth: 200, isResizable: true },
        {
            key: 'ProjectManager',
            name: 'Project Manager',
            fieldName: 'ProjectManager',
            minWidth: 180,
            maxWidth: 250,
            isResizable: true,
            onRender: (item) => (
                item.ProjectManager ? (
                    <Persona
                        text={item.ProjectManagerTitle}
                        size={PersonaSize.size32}
                        imageUrl={item.ProjectManagerPhotoUrl}
                    />
                ) : ''
            )
        },
        { key: 'Status', name: 'Status', fieldName: 'Status', minWidth: 80, maxWidth: 120, isResizable: true },
        { key: 'StartDate', name: 'Start Date', fieldName: 'StartDate', onRender: (item) => formatDate(item.StartDate), minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'EndDate', name: 'End Date', fieldName: 'EndDate', onRender: (item) => formatDate(item.EndDate), minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'Notes', name: 'Notes', fieldName: 'Notes', minWidth: 120, maxWidth: 200, isResizable: true },
        { key: 'CreatedBy', name: 'Created By', fieldName: 'CreatedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.CreatedBy || '' },
        { key: 'Created', name: 'Created', fieldName: 'Created', onRender: (item) => formatDate(item.Created), minWidth: 100, maxWidth: 140, isResizable: true },
        { key: 'ModifiedBy', name: 'Modified By', fieldName: 'ModifiedBy', minWidth: 120, maxWidth: 200, isResizable: true, onRender: (item) => item.ModifiedBy || '' },
        { key: 'Modified', name: 'Modified', fieldName: 'Modified', onRender: (item) => formatDate(item.Modified), minWidth: 100, maxWidth: 140, isResizable: true },
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
                                setEditProjectId(null);
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
                                setIsModalOpen(true);
                            },
                            disabled: false,
                            color: '#5a2a6b'
                        },
                        selectedItems.length === 1 ? {
                            key: 'editProject',
                            text: 'Edit Project',
                            iconName: 'Edit',
                            onClick: () => {
                                const item = selectedItems[0];
                                setEditProjectId(item.Id || null);
                                setNewProject({
                                    ProjectCode: item.ProjectCode || '',
                                    Title: item.Title || '',
                                    Customer: item.Customer || '',
                                    ProjectManagerId: item.ProjectManagerId,
                                    ProjectManagerTitle: item.ProjectManagerTitle || '',
                                    Status: item.Status || '',
                                    StartDate: item.StartDate || '',
                                    EndDate: item.EndDate || '',
                                    Notes: item.Notes || ''
                                });
                                setIsModalOpen(true);
                            },
                            disabled: false,
                            color: '#5a2a6b',
                        } : null,
                        selectedItems.length > 0 ? {
                            key: 'deleteProject',
                            text: 'Delete Project',
                            iconName: 'Delete',
                            onClick: () => setShowDeleteConfirm(true),
                            disabled: false,
                            color: '#a4262c',
                        } : null
                    ].filter((b): b is NonNullable<typeof b> => b !== null)
                }
            />
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    <DetailsList
                        items={items}
                        columns={columns}
                        setKey="set"
                        selectionMode={2}
                        selection={selectionRef.current!}
                        selectionPreservedOnEmptyClick={true}
                        layoutMode={DetailsListLayoutMode.justified}
                        constrainMode={ConstrainMode.horizontalConstrained}
                    />
                </div>
            )}
            {isModalOpen && (
                <ModalContainer
                    isOpen={isModalOpen}
                    title={editProjectId ? "Modifica Project" : "Crea nuovo Project"}
                    onSave={handleSaveProject}
                    onCancel={() => { setIsModalOpen(false); setEditProjectId(null); }}
                    saving={saving}
                >
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
                        context={context}
                        titleText="Project Manager"
                        personSelectionLimit={1}
                        showtooltip={true}
                        required={true}
                        onChange={items => {
                            setNewProject(p => ({
                                ...p,
                                ProjectManagerId: items[0] ? Number(items[0]) : undefined
                            }));
                        }}
                        principalTypes={[1]}
                        resolveDelay={300}
                        ensureUser={true}
                        showHiddenInUI={false}
                        suggestionsLimit={15}
                        defaultSelectedUsers={editProjectId && newProject.ProjectManagerId && newProject.ProjectManagerTitle ? [{
                            text: newProject.ProjectManagerTitle,
                            secondaryText: '',
                            id: String(newProject.ProjectManagerId)
                        }] : []}
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
                        isRequired={true}
                    />
                    <DatePicker
                        label="End Date"
                        value={newProject.EndDate ? new Date(newProject.EndDate) : undefined}
                        onSelectDate={date => setNewProject(p => ({ ...p, EndDate: date ? date.toISOString().substring(0, 10) : '' }))}
                        placeholder="DD-MM-YYYY"
                        formatDate={d => d ? d.toLocaleDateString() : ''}
                        isRequired={true}
                    />
                    <TextField label="Notes" multiline rows={3} value={newProject.Notes} onChange={(_, v) => setNewProject(p => ({ ...p, Notes: v || '' }))} />
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