import * as React from 'react';
import { IDocumentMetadataProps, IDocumentMetadataFormProps } from './IDocumentMetadataFormProps';
import { TextField, Stack, Dropdown, IDropdownOption, MessageBar } from '@fluentui/react';
import { ModalContainer } from '../ModalContainer/ModalContainer';
import { PeoplePicker } from '../PeoplePicker/PeoplePicker';
import { DocumentsService, UsersService } from '../../../services/SharePointService';




export const DocumentMetadataForm: React.FC<IDocumentMetadataFormProps> = ({
  initialValues = {},
  onSubmit,
  onCancel,
  saving,
  context,
  isOpen = true,
  title = 'Dettagli Documento',
  width = '480px',
  saveText = 'Salva',
  cancelText = 'Annulla'
}) => {
  const [values, setValues] = React.useState<IDocumentMetadataProps>({
    DocumentCode: initialValues.DocumentCode || '',
    Title: initialValues.Title || '',
    Revision: initialValues.Revision || '',
    Status: initialValues.Status || '',
    IssuePurpose: initialValues.IssuePurpose || '',
    ApprovalCode: initialValues.ApprovalCode || '',
    AssignedToId: initialValues.AssignedToId,
    Notes: initialValues.Notes || '',
  });

  const [statusOptions, setStatusOptions] = React.useState<IDropdownOption[]>([]);
  const [issuePurposeOptions, setIssuePurposeOptions] = React.useState<IDropdownOption[]>([]);
  const [approvalCodeOptions, setApprovalCodeOptions] = React.useState<IDropdownOption[]>([]);
  const [usersList, setUsersList] = React.useState<any[]>([]);

  React.useEffect(() => {
    // Load Status choices
    const fetchStatusChoices = async () => {
      try {
        const documentsService = new DocumentsService(context);
        const choices = await documentsService.getStatusChoices();
        setStatusOptions(choices.map((c: string) => ({ key: c, text: c })));
      } catch {
        setStatusOptions([]);
      }
    };
    // Load IssuePurpose choices
    const fetchIssuePurposeChoices = async () => {
      try {
        const documentsService = new DocumentsService(context);
        const choices = await documentsService.getIssuePurposeChoices();
        setIssuePurposeOptions(choices.map((c: string) => ({ key: c, text: c })));
      } catch {
        setIssuePurposeOptions([]);
      }
    };
    // Load ApprovalCode choices
    const fetchApprovalCodeChoices = async () => {
      try {
        const documentsService = new DocumentsService(context);
        const choices = await documentsService.getApprovalCodeChoices();
        setApprovalCodeOptions(choices.map((c: string) => ({ key: c, text: c })));
      } catch {
        setApprovalCodeOptions([]);
      }
    };
    // Load users for PeoplePicker
    const fetchUsers = async () => {
      try {
        const usersService = new UsersService(context);
        const users = await usersService.getUsers();
        setUsersList(users);
      } catch {
        setUsersList([]);
      }
    };
    void fetchStatusChoices();
    void fetchIssuePurposeChoices();
    void fetchApprovalCodeChoices();
    void fetchUsers();
  }, [context]);

  const handleChange = (field: keyof IDocumentMetadataProps, value: any) => {
    setValues(v => ({ ...v, [field]: value }));
  };

  // AssignedTo PeoplePicker handler
  const handleAssignedToChange = (users: any[]) => {
    handleChange('AssignedToId', users && users.length > 0 ? Number(users[0].id) : undefined);
  };

  // Validation state (like Projects.tsx)
  const [touched, setTouched] = React.useState(false);
  const [formError, setFormError] = React.useState<string | null>(null);
  const isValid =
    values.DocumentCode.trim() !== '' &&
    values.Title.trim() !== '' &&
    values.Revision.trim() !== '' &&
    values.Status.trim() !== '' &&
    values.IssuePurpose.trim() !== '' &&
    values.ApprovalCode.trim() !== '' &&
    !!values.AssignedToId;

  const handleSave = () => {
    setTouched(true);
    if (!isValid) {
      setFormError('Compila tutti i campi obbligatori prima di salvare.');
      return;
    }
    setFormError(null);
    onSubmit(values);
  };

  return (
    <ModalContainer
      isOpen={isOpen}
      title={title}
      onSave={handleSave}
      onCancel={onCancel}
      saving={saving}
      saveText={saveText}
      cancelText={cancelText}
      width={width}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        {touched && formError && (
          <MessageBar messageBarType={2}>{formError}</MessageBar>
        )}
        <TextField
          label="Document Code"
          value={values.DocumentCode}
          onChange={(_, v) => handleChange('DocumentCode', v || '')}
          required
        />
        <TextField
          label="Title"
          value={values.Title}
          onChange={(_, v) => handleChange('Title', v || '')}
          required
        />
        <TextField
          label="Revision"
          value={values.Revision}
          onChange={(_, v) => handleChange('Revision', v || '')}
          required
        />
        <Dropdown
          label="Status"
          options={statusOptions}
          selectedKey={values.Status}
          onChange={(_, o) => handleChange('Status', o?.key as string || '')}
          required
        />
        <Dropdown
          label="Issue Purpose"
          options={issuePurposeOptions}
          selectedKey={values.IssuePurpose}
          onChange={(_, o) => handleChange('IssuePurpose', o?.key as string || '')}
          required
        />
        <Dropdown
          label="Approval Code"
          options={approvalCodeOptions}
          selectedKey={values.ApprovalCode}
          onChange={(_, o) => handleChange('ApprovalCode', o?.key as string || '')}
          required
        />
        <PeoplePicker
          context={context}
          selectedUserIds={values.AssignedToId ? [values.AssignedToId] : []}
          onChange={handleAssignedToChange}
          required
          label="Assigned To"
          itemLimit={1}
          loadUsers={async () => {
            // Ensure the selected user is present in the list for correct display
            if (values.AssignedToId && usersList.every(u => String(u.id) !== String(values.AssignedToId))) {
              // Optionally fetch the user by ID and add to usersList if needed
              // For now, just return usersList as is
              return usersList;
            }
            return usersList;
          }}
        />
        <TextField
          label="Notes"
          multiline
          rows={3}
          value={values.Notes}
          onChange={(_, v) => handleChange('Notes', v || '')}
        />
      </Stack>
    </ModalContainer>
  );
};
