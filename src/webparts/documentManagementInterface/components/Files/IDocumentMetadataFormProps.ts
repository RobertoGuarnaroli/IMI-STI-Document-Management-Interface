import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentMetadataProps {
  DocumentCode: string;
  Title: string;
  Revision: string;
  Status: string;
  IssuePurpose: string;
  ApprovalCode: string;
  AssignedToId?: number;
  Notes?: string;
}

export interface IDocumentMetadataFormProps {
  initialValues?: Partial<IDocumentMetadataProps>;
  onSubmit: (values: IDocumentMetadataProps) => void;
  onCancel: () => void;
  saving?: boolean;
  context: WebPartContext;
  isOpen?: boolean;
  title?: string;
  width?: string;
  saveText?: string;
  cancelText?: string;
}