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

export const Projects: React.FC<IProjectsProps> = ({ context }) => {
    const [items, setItems] = React.useState<IProjectItem[]>([]);
    const [loading, setLoading] = React.useState(true);
    const [isModalOpen, setIsModalOpen] = React.useState(false);
    const [newProject, setNewProject] = React.useState({
        ProjectCode: '',
        Title: '',
        Customer: '',
        ProjectManagerId: undefined as number | undefined,
        Status: '',
        StartDate: '',
        EndDate: '',
        Notes: ''
    });

    const [statusOptions, setStatusOptions] = React.useState<IDropdownOption[]>([]);
    const [saving, setSaving] = React.useState(false);
    const [dateError, setDateError] = React.useState<string | null>(null);
    const [showError, setShowError] = React.useState(false);
    const [formError, setFormError] = React.useState<string | null>(null);

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
            setIsModalOpen(false);
            setNewProject({ ProjectCode: '', Title: '', Customer: '', ProjectManagerId: undefined, Status: '', StartDate: '', EndDate: '', Notes: '' });
            setLoading(true);
            // Ricarica la lista
            const data = await service.getProjects();
            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
            const mapped: IProjectItem[] = data.map((item) => ({
                ProjectCode: item.ProjectCode || '',
                Title: item.Title || '',
                Customer: item.Customer || '',
                ProjectManager: item.ProjectManager?.Title || '',
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
        } catch {
            // TODO: gestione errore
        } finally {
            setSaving(false);
            setLoading(false);
        }
    };

    const handleDeleteProject = async (id: number) => {
        setSaving(true);
        try {
            const service = new ProjectsService(context);
            await service.deleteProject(id);
            // Aggiorna la lista dopo l'eliminazione
            const data = await service.getProjects();
            const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, '').trim();
            const mapped: IProjectItem[] = data.map((item) => ({
                ProjectCode: item.ProjectCode || '',
                Title: item.Title || '',
                Customer: item.Customer || '',
                ProjectManager: item.ProjectManager?.Title || '',
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
                const stripHtml = (html: string) => html.replace(/<[^>]+>/g, '').trim();
                const mapped: IProjectItem[] = data.map((item) => ({
                    Id: item.Id,
                    ProjectCode: item.ProjectCode || '',
                    Title: item.Title || '',
                    Customer: item.Customer || '',
                    ProjectManager: item.ProjectManager?.Title || '',
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
        const fetchStatusOptions = async () => {
            const service = new ProjectsService(context);
            const choices = await service.getStatusChoices();
            setStatusOptions(choices.map(choice => ({ key: choice, text: choice })));
        };
        void fetchStatusOptions();
    }, [context]);

    // Imposta lo stato predefinito e unico per nuovo progetto
    React.useEffect(() => {
        if (isModalOpen) {
            setNewProject(p => ({ ...p, Status: 'Active' }));
            setStatusOptions([{ key: 'Active', text: 'Active' }]);
        }
    }, [isModalOpen]);

    const formatDate = (dateStr?: string) => {
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
                        text={item.ProjectManager}
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
                buttons={[
                    {
                        key: 'addProject',
                        text: 'New Project',
                        iconName: 'Add',
                        onClick: () => setIsModalOpen(true),
                        disabled: false,
                        color: '#5a2a6b'
                    },
                    {
                        key: 'deleteProject',
                        text: 'Delete Project',
                        iconName: 'Delete',
                        onClick: () => {
                            const selectedRows = items.filter(item => item.isSelected);
                            if (selectedRows.length === 0) {
                                setFormError('Seleziona almeno un progetto da eliminare.');
                                setShowError(true);
                                return;
                            }
                            Promise.all(selectedRows.map(async (item) => {
                                if (item.Id) {
                                    await handleDeleteProject(item.Id);
                                }
                            })).then(() => {
                                setShowError(false);
                            }).catch(() => {
                                setFormError('Errore durante l\'eliminazione dei progetti.');
                                setShowError(true);
                            });
                        },
                        disabled: false,
                        color: '#a4262c',
                        visible: items.some(item => item.isSelected)
                    }
                ]}
            />
            {loading ? (
                <LoadingSpinner />
            ) : (
                <div className={styles.listContainer}>
                    <DetailsList
                        items={items}
                        columns={columns}
                        layoutMode={DetailsListLayoutMode.justified}
                        constrainMode={ConstrainMode.horizontalConstrained}                       
                    />
                </div>
            )}
            {isModalOpen && (
                <ModalContainer
                    isOpen={isModalOpen}
                    title="Crea nuovo Project"
                    onSave={handleSaveProject}
                    onCancel={() => setIsModalOpen(false)}
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
                        context={context as any}
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
                        defaultSelectedUsers={[]}
                        disabled={false}
                        label="Project Manager"
                        placeholder="Seleziona il Project Manager"
                        itemLimit={1}
                        loadUsers={async (context) => {
                            const service = new UsersService(context as any);
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
        </div>
    );
}