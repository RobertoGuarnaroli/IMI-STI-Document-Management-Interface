import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPersonaProps } from '@fluentui/react/lib/Persona';
import { PrincipalType } from '@pnp/sp';

export interface IGenericPeoplePickerProps {
  titleText?: string;
  personSelectionLimit?: number;
  showtooltip?: boolean;
  context: WebPartContext;
  selectedUserIds?: number[];
  onChange: (userIds: number[]) => void;
  principalTypes?: PrincipalType[];
  resolveDelay?: number;
  ensureUser?: boolean;
  showHiddenInUI?: boolean;
  suggestionsLimit?: number;
  defaultSelectedUsers?: IPersonaProps[];
  required?: boolean;
  disabled?: boolean;
  label?: string;
  itemLimit?: number;
  placeholder?: string;
  loadUsers: (context: WebPartContext) => Promise<IPersonaProps[]>;
}

export interface IGenericPeoplePickerState {
  usersList: IPersonaProps[];
  selectedUsers: IPersonaProps[];
}
