export interface IDocumentsProps {
 context: any;
}

export interface IDocumentItem {
  DocumentCode: string;
  Title: string;
  Revision: string;
  Status: string;
  IssuePurpose: string;
  ApprovalCode: string;
  SentDate: string;
  ExpectedReturnDate: string;
  ActualReturnDate: string;
  TurnaroundDays: number;
  DaysLate: number;
  AssignedTo?: {
    Id: number;
    Title: string;
    Picture?: string;
  };
  Notes: string;
  Modified: string;
  Created: string;
  CreatedBy: string;
  ModifiedBy: string;
}
