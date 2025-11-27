import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentHistoryProps {
  context: WebPartContext;
}

export interface IDocumentHistoryItem {
  DocumentId: number;
  DocumentCode: string;
  Revision: string;
  Action: string;
  ActionDate: string;
  PerformedBy?: {
    Id: number;
    Title: string;
    EMail: string;
    Picture?: string;
  };
  Status: string;
  ApprovalCode: string;
}
