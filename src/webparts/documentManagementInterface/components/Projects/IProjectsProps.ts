export interface IProjectsProps {
	context: any;
}

export interface IProjectItem{
    key: number;
    Id?: number;
	ProjectCode?: string;
    Title?: string;
    Customer?: string;
    ProjectManager:{
        Id: number;
        Title: string;
        Picture?: string;
    };
    Status?: string;
    StartDate?: string;
    EndDate?: string;
    Notes?: string;
    Modified?: string;
    Created?: string;
    CreatedBy?: string;
    ModifiedBy?: string;
    context: unknown;
    isSelected?: boolean;
}