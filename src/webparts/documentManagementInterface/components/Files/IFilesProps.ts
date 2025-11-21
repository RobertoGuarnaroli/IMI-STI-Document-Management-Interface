export interface IFilesProps{
    context: any;
}
export interface IFileItem {
  Created?: string;
  Modified?: string;
  FileName?: string;
  CheckedOutTo?: {
    Id: number;
    Title: string;
  };
  CreatedBy?: {
    Id: number;
    Title: string;
  };
  ModifiedBy?: {
    Id: number;
    Title: string;
  };
}