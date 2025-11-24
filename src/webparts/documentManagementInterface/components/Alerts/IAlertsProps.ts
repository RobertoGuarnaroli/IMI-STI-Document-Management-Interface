import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAlertsProps {
  context: WebPartContext;
}

export interface IAlertItem {
  Id: number;
  ProjectCode?: {
    Id: number;
    ProjectCode: string;
  };
  DocumentId?: {
    Id: number;
    DocumentCode: string;
  };
  AlertType: string;
  Priority: string;
  DaysOverdue: number;
  ExpectedDate: string;
  Message: string;
  AssignedTo?: {
    Id: number;
    Title: string;
    EMail: string;
  };
  IsResolved: boolean;
  ResolvedDate?: string;
}
