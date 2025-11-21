import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPersonaProps } from "@fluentui/react/lib/Persona";

export type Project = Record<string, any>;

// ====================
// 📌 USERS SERVICE
// ====================
export class UsersService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    /**
     * Recupera gli utenti SharePoint (solo utenti, non gruppi)
     * in formato IPersonaProps per il People Picker.
     */
    public async getUsers(): Promise<IPersonaProps[]> {
        try {
            const users = await this.sp.web.siteUsers
                .filter("PrincipalType eq 1")   // 1 = utenti, no gruppi
                .top(200)();                    // aumentabile se necessario

            return users.map((u: any) => ({
                text: u.Title,
                secondaryText: u.Email,
                id: u.Id.toString(),
            }));
        } catch (error) {
            console.error("Errore nel recupero degli utenti:", error);
            return [];
        }
    }
}


/*===================================================================================
    PROJECTS SERVICE
====================================================================================*/

export class ProjectsService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getProjects(): Promise<Project[]> {
        try {
            const items = await this.sp.web.lists.getByTitle("Projects")
                .items
                .select(
                    "*",
                    "ProjectManager/Id",
                    "ProjectManager/Title",
                    "Author/Id",
                    "Author/Title",
                    "Editor/Id",
                    "Editor/Title"
                )
                .expand("ProjectManager", "Author", "Editor")();

            console.log('Raw SharePoint items:', items);
            return items as Project[];
        } catch (error) {
            console.error("Errore nel recupero dei progetti:", error);
            throw error;
        }
    }

    public async addProject(project: {
        ProjectCode: string;
        Title: string;
        Customer: string;
        ProjectManagerId?: number;
        Status: string;
        StartDate: string;
        EndDate: string;
        Notes: string;
    }): Promise<void> {
        try {
            await this.sp.web.lists.getByTitle("Projects").items.add({
                ProjectCode: project.ProjectCode,
                Title: project.Title,
                Customer: project.Customer,
                ProjectManagerId: project.ProjectManagerId,
                Status: project.Status,
                StartDate: project.StartDate,
                EndDate: project.EndDate,
                Notes: project.Notes
            });
        } catch (error) {
            console.error("Errore nell'aggiunta del progetto:", error);
            throw error;
        }
    }

    public async getStatusChoices(): Promise<string[]> {
        try {
            const field = await this.sp.web.lists.getByTitle("Projects").fields.getByInternalNameOrTitle("Status")();
            return field.Choices as string[];
        } catch (error) {
            console.error("Errore nel recupero delle opzioni Status:", error);
            return [];
        }
    }
}


/*===================================================================================
    DOCUMENTS SERVICE
====================================================================================*/

export class DocumentsService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getDocuments(): Promise<any[]> {
        try {
            const items = await this.sp.web.lists.getByTitle("Documents")
                .items
                .select(
                    "*",
                    "AssignedTo/Id",
                    "AssignedTo/Title",
                    "Author/Id",
                    "Author/Title",
                    "Editor/Id",
                    "Editor/Title"
                )
                .expand("AssignedTo", "Author", "Editor")();

            console.log('Raw SharePoint items:', items);
            return items;
        } catch (error) {
            console.error("Errore nel recupero dei documenti:", error);
            throw error;
        }
    }
}


/*===================================================================================
    TRANSMITTALS SERVICE
====================================================================================*/

export class TransmittalsService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getTransmittals(): Promise<any[]> {
        try {
            const items = await this.sp.web.lists.getByTitle("Transmittals")
                .items
                .select(
                    "*",
                    "ProjectCode/Id",
                    "ProjectCode/ProjectCode",
                    "ProjectCode/Title",
                    "Author/Id",
                    "Author/Title",
                    "Editor/Id",
                    "Editor/Title"
                )
                .expand("ProjectCode", "Author", "Editor")();

            console.log('Raw SharePoint items:', items);
            return items;
        } catch (error) {
            console.error("Errore nel recupero dei transmittals:", error);
            throw error;
        }
    }
}


/*===================================================================================
    DISTRIBUTION LISTS SERVICE
====================================================================================*/

export class DistributionListsService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getDistributionLists(): Promise<any[]> {
        try {
            const items = await this.sp.web.lists.getByTitle("DistributionList")
                .items
                .select(
                    "*",
                    "ProjectCode/Id",
                    "ProjectCode/ProjectCode",
                    "ProjectCode/Title",
                    "Author/Id",
                    "Author/Title",
                    "Editor/Id",
                    "Editor/Title"
                )
                .expand("ProjectCode", "Author", "Editor")();

            console.log('Raw SharePoint items:', items);
            return items;
        } catch (error) {
            console.error("Errore nel recupero delle distribution lists:", error);
            throw error;
        }
    }
}


/*===================================================================================
    FILES SERVICE
====================================================================================*/

export class FilesService {
    private sp: SPFI;

    constructor(context: WebPartContext) {
        this.sp = spfi().using(SPFx(context));
    }

    public async getFiles(): Promise<any[]> {
        try {
            const items = await this.sp.web.lists.getByTitle("Files")
                .items
                .select(
                    "*",
                    "FileLeafRef",
                    "FileRef",
                    "Author/Id",
                    "Author/Title",
                    "Editor/Id",
                    "Editor/Title",
                    "CheckoutUser/Id",
                    "CheckoutUser/Title"
                )
                .expand("Author", "Editor", "CheckoutUser")();

            console.log('Raw SharePoint items:', items);
            return items;
        } catch (error) {
            console.error("Errore nel recupero dei files:", error);
            throw error;
        }
    }
}
