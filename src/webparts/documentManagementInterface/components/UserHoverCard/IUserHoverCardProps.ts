export interface IUserHoverCardProps {
  user?: {
    id: string;
    displayName: string;
    mail?: string;
    jobTitle?: string;
    department?: string;
    officeLocation?: string;
    businessPhones?: string[];
    mobilePhone?: string;
    pictureUrl?: string;
  };
  loading?: boolean;
}
