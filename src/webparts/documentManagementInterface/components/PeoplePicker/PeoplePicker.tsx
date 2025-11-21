import * as React from 'react';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import {
  IBasePickerSuggestionsProps,
  NormalPeoplePicker,
  ValidationState,
} from '@fluentui/react/lib/Pickers';
import { IGenericPeoplePickerProps } from './IPeoplePicketProps';

const suggestionProps: IBasePickerSuggestionsProps = {
  suggestionsHeaderText: 'Utenti suggeriti',
  mostRecentlyUsedHeaderText: 'Contatti recenti',
  noResultsFoundText: 'Nessun risultato trovato',
  loadingText: 'Caricamento...',
  showRemoveButtons: true,
  suggestionsAvailableAlertText: 'Suggerimenti disponibili',
  suggestionsContainerAriaLabel: 'Contatti suggeriti',
};

export const PeoplePicker: React.FC<IGenericPeoplePickerProps> = ({
  context,
  selectedUserIds = [],
  onChange,
  required = false,
  disabled = false,
  label,
  itemLimit = 1,
  placeholder,
  loadUsers, // funzione esterna per ottenere gli utenti
}) => {

  const [usersList, setUsersList] = React.useState<IPersonaProps[]>([]);
  const [selectedUsers, setSelectedUsers] = React.useState<IPersonaProps[]>([]);

  /** -------------------------
   * CARICAMENTO UTENTI ESTERNO
   * -------------------------*/
  React.useEffect(() => {
    const fetchUsers = async () => {
      try {
        const users = await loadUsers(context);  // usa funzione esterna
        setUsersList(users);
      } catch (error) {
        console.error("Errore caricamento utenti:", error);
      }
    };

    void fetchUsers();
  }, [context, loadUsers]);

  /** -------------------------
   * SYNC utenti selezionati da ID
   * -------------------------*/
  React.useEffect(() => {
    if (selectedUserIds.length > 0 && usersList.length > 0) {
      const selected = usersList.filter(user =>
        selectedUserIds.includes(Number(user.id))
      );
      setSelectedUsers(selected);
    }
  }, [selectedUserIds, usersList]);

  /** -------------------------
   * FILTRO SUGGERIMENTI
   * -------------------------*/
  const onFilterChanged = (
    filterText: string,
    currentPersonas: IPersonaProps[]
  ): IPersonaProps[] => {
    if (!filterText) return [];

    const filtered = usersList.filter(user =>
      user.text?.toLowerCase().includes(filterText.toLowerCase()) ||
      user.secondaryText?.toLowerCase().includes(filterText.toLowerCase())
    );

    return removeDuplicates(filtered, currentPersonas).slice(0, 10);
  };

  /** -------------------------
   * SELEZIONE UTENTI
   * -------------------------*/
  const onSelectionChanged = (items?: IPersonaProps[]) => {
    const selected = items || [];
    setSelectedUsers(selected);

    const userIds = selected.map(item => Number(item.id));
    onChange(userIds);
  };

  return (
    <NormalPeoplePicker
      label={label}
      onResolveSuggestions={onFilterChanged}
      getTextFromItem={(p) => p.text || ""}
      pickerSuggestionsProps={suggestionProps}
      className={'ms-PeoplePicker'}
      onValidateInput={(input) =>
        input.length > 2 ? ValidationState.valid : ValidationState.invalid
      }
      selectionAriaLabel={label ? `${label} selezionato` : 'Utente selezionato'}
      removeButtonAriaLabel="Rimuovi"
      inputProps={{
        'aria-label': label || 'People Picker',
        placeholder: placeholder,
      }}
      resolveDelay={300}
      disabled={disabled}
      required={required}
      onChange={onSelectionChanged}
      selectedItems={selectedUsers}
      itemLimit={itemLimit}
    />
  );
};

/* -------------------------
   HELPER FUNCTIONS
--------------------------*/
function removeDuplicates(personas: IPersonaProps[], selected: IPersonaProps[]) {
  return personas.filter(persona => !listContainsPersona(persona, selected));
}

function listContainsPersona(persona: IPersonaProps, personas: IPersonaProps[]) {
  return personas.some(item => item.id === persona.id);
}
