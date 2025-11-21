export interface IDistributionListsProps {
    context: any;   
}

export interface IDistributionListsItem {
  ProjectCode?: number;
  ContactEmail?: string;
  Company?: string;
  Role?: string;
  ReceiveTransmittals?: boolean;
  ReceiveNotifications?: boolean;
  ReceiveReminders?: boolean;
  IsActive?: boolean;
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
