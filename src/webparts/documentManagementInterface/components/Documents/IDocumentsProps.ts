export interface IDocumentsProps {
 context: any;
}

export interface IDocumentItem {
 DocumentCode?: string;
  Title?: string;
  Revision?: string;
  Status?: string;
  IssuePurpose?: string;
  ApprovalCode?: string;
  SentDate?: string;
  ExpectedReturnDate?: string;
  ActualReturnDate?: string;
  TurnaroundDays?: number;
  DaysLate?: number;
  AssignedTo?: {
    Id: number;
    Title: string;
  };
  Notes?: string;
  FileId?: number;
  Modified?: string;
  Created?: string;
  CreatedBy?: {
    Id: number;
    Title: string;
  };
  ModifiedBy?: {
    Id: number;
    Title: string;
  };
}