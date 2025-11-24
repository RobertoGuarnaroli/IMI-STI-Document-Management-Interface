export interface IProjectsProps {
	context: any;
}

export interface IProjectItem{
    key: number;
    Id?: number;
	ProjectCode?: string;
    Title?: string;
    Customer?: string;
    ProjectManagerId?: number;
    ProjectManagerTitle?: string;
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