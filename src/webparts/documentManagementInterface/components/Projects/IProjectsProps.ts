export interface IProjectsProps {
	context: any;
}

export interface IProjectItem{
	ProjectCode?: string;
    Title?: string;
    Customer?: string;
    ProjectManager?: string;
    Status?: string;
    StartDate?: string;
    EndDate?: string;
    Notes?: string;
    Modified?: string;
    Created?: string;
    CreatedBy?: string;
    ModifiedBy?: string;
    context: unknown;
}