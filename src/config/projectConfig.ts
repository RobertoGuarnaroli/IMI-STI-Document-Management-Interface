
    export interface ProjectConfig {
        siteUrl: string;
        lists: Array<{
            id: string;
            name: string;
            description?: string;
            icon?: string;
        }>;
        map<T>(fn: (list: { id: string; name: string; description?: string }) => T): T[];
    }

const projectConfig: ProjectConfig = {
    siteUrl: "https://pplutech.sharepoint.com/sites/IMI-STI",
    lists: [
        {
            id: "projects",
            name: "Projects",
            description: "List of available projects",
            icon: "FabricFolder"
        },
        {
            id: "documents",
            name: "Documents",
            description: "List of available documents",
            icon: "Document"
        },
        {
            id: "transmittals",
            name: "Transmittals",
            description: "List of available transmittals",
            icon: "Send"
        },
        {
            id: "distributionLists",
            name: "Distribution Lists",
            description: "List of available distribution lists",
            icon: "People"
        },
        {
            id: "documentHistory",
            name: "Document History",
            description: "List of document history",
            icon: "History"
        },
        {
            id: "alerts",
            name: "Alerts",
            description: "List of available alerts",
            icon: "Ringer"
        },
        {
            id: "files",
            name: "Files",
            description: "Document storage",
            icon: "OpenFile"
        }
    ],
    map(fn) {
        return this.lists.map(fn);
    }
};

export default projectConfig;