export interface ITransmittalsProps {
    context: any;
}

export interface ITransmittalItem {
  TransmittalNumber?: string;
  ProjectCode?: number;
  ProjectCodeTitle?: string;
  TransmittalType?: string;
  SentDate?: string;
  ReceivedDate?: string;
  RecipientEmail?: string;
  RecipientName?: string;
  SenderEmail?: string;
  SenderName?: string;
  Subject?: string;
  Notes?: string;
  Documents?: string;
  Status?: string;
  Modified?: string;
  Created?: string;
  CreatedBy?: {
    Id: number;
    Title: string;
    EMail?: string;
    Picture?: string;
  };
  ModifiedBy?: {
    Id: number;
    Title: string;
    EMail?: string;
    Picture?: string;
  };
}
