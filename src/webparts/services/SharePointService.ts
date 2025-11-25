
import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
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
        .filter("PrincipalType eq 1") // 1 = utenti, no gruppi
        .top(5000)(); // aumentabile se necessario

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
    USER PROFILE SERVICE
====================================================================================*/

export class UserProfileService {
  private sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Recupera la URL dell'immagine profilo di un utente dato l'ID o l'email
   * @param userIdOrEmail ID numerico o email dell'utente
   * @returns URL dell'immagine profilo o stringa vuota
   */
  public async getUserProfilePicture(userIdOrEmail: number | string): Promise<string> {
    try {
      let loginName: string | undefined;
      if (typeof userIdOrEmail === 'number') {
        // Recupera loginName da ID
        const user = await this.sp.web.siteUsers.getById(userIdOrEmail)();
        loginName = user.LoginName;
      } else {
        // Recupera loginName da email
        const users = await this.sp.web.siteUsers.filter(`Email eq '${userIdOrEmail}'`)();
        if (users && users.length > 0) {
          loginName = users[0].LoginName;
        }
      }
      if (!loginName) return '';
      // Ottieni la URL della foto profilo
      const photoUrl = `${this.sp.web.toUrl()}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(loginName)}`;
      return photoUrl;
    } catch (error) {
      console.error('Errore nel recupero della foto profilo:', error);
      return '';
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
      const items = await this.sp.web.lists
        .getByTitle("Projects")
        .items.select(
          "*",
          "ProjectManager/Id",
          "ProjectManager/Title",
          "Author/Id",
          "Author/Title",
          "Editor/Id",
          "Editor/Title"
        )
        .expand("ProjectManager", "Author", "Editor")();

      console.log("Raw SharePoint items:", items);
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
        Notes: project.Notes,
      });
    } catch (error) {
      console.error("Errore nell'aggiunta del progetto:", error);
      throw error;
    }
  }

  /**
   * Aggiorna un progetto esistente dato l'ID e i nuovi dati
   */
  public async updateProject(
    itemId: number,
    project: {
      ProjectCode: string;
      Title: string;
      Customer: string;
      ProjectManagerId?: number;
      Status: string;
      StartDate: string;
      EndDate: string;
      Notes: string;
    }
  ): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle("Projects")
        .items.getById(itemId)
        .update({
          ProjectCode: project.ProjectCode,
          Title: project.Title,
          Customer: project.Customer,
          ProjectManagerId: project.ProjectManagerId,
          Status: project.Status,
          StartDate: project.StartDate,
          EndDate: project.EndDate,
          Notes: project.Notes,
        });
    } catch (error) {
      console.error("Errore nell'aggiornamento del progetto:", error);
      throw error;
    }
  }

  public async getStatusChoices(): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle("Projects")
        .fields.getByInternalNameOrTitle("Status")();
      return field.Choices as string[];
    } catch (error) {
      console.error("Errore nel recupero delle opzioni Status:", error);
      return [];
    }
  }

  public async deleteProject(itemId: number): Promise<void> {
    try {
      await this.sp.web.lists
        .getByTitle("Projects")
        .items.getById(itemId)
        .delete();
    } catch (error) {
      console.error("Errore nella cancellazione del progetto:", error);
      throw error;
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
      const items = await this.sp.web.lists
        .getByTitle("Documents")
        .items.select(
          "*",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "Author/Id",
          "Author/Title",
          "Editor/Id",
          "Editor/Title"
        )
        .expand("AssignedTo", "Author", "Editor")();

      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero dei documenti:", error);
      throw error;
    }
  }

  /**
     * Inserisce un nuovo documento nella lista Documents
     */
    public async createDocument(document: {
      DocumentCode: string;
      Title: string;
      Revision: string;
      Status: string;
      IssuePurpose: string;
      ApprovalCode: string;
      SentDate?: string;
      ExpectedReturnDate?: string;
      ActualReturnDate?: string;
      TurnaroundDays?: number;
      DaysLate?: number;
      AssignedToId?: number;
      Notes?: string;
    }): Promise<void> {
      try {
        await this.sp.web.lists.getByTitle("Documents").items.add({
          DocumentCode: document.DocumentCode,
          Title: document.Title,
          Revision: document.Revision,
          Status: document.Status,
          IssuePurpose: document.IssuePurpose,
          ApprovalCode: document.ApprovalCode,
          SentDate: document.SentDate,
          ExpectedReturnDate: document.ExpectedReturnDate,
          ActualReturnDate: document.ActualReturnDate,
          TurnaroundDays: document.TurnaroundDays,
          DaysLate: document.DaysLate,
          AssignedToId: document.AssignedToId,
          Notes: document.Notes,
        });
      } catch (error) {
        console.error("Errore durante l'inserimento del documento:", error);
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
      const items = await this.sp.web.lists
        .getByTitle("Transmittals")
        .items.select(
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

      console.log("Raw SharePoint items:", items);
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
      const items = await this.sp.web.lists
        .getByTitle("DistributionList")
        .items.select(
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

      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero delle distribution lists:", error);
      throw error;
    }
  }
}

/*===================================================================================
    DOCUMENT HISTORY SERVICE
====================================================================================*/

export class DocumentHistoryService {
  private sp: SPFI;
  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }
  public async getDocumentHistory(): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("DocumentHistory")
        .items.select(
          "*",
          "DocumentId/Id",
          "DocumentId/DocumentCode",
          "DocumentId/Revision",
          "PerformedBy/Id",
          "PerformedBy/Title",
          "PerformedBy/EMail",
          "Author/Id",
          "Author/Title",
          "Editor/Id",
          "Editor/Title"
        )
        .expand("DocumentId", "PerformedBy", "Author", "Editor")();
      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero della cronologia documenti:", error);
      throw error;
    }
  }
}

/*===================================================================================
    ALERTS SERVICE
====================================================================================*/

export class AlertsService {
  private sp: SPFI;
  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }
  public async getAlerts(): Promise<any[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("Alerts")
        .items.select(
          "*",
          "ProjectCode/Id",
          "ProjectCode/ProjectCode",
          "DocumentId/Id",
          "DocumentId/DocumentCode",
          "AlertType",
          "Priority",
          "DaysOverdue",
          "ExpectedDate",
          "Message",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "IsResolved",
          "ResolvedDate"
        )
        .expand("ProjectCode", "DocumentId", "AssignedTo")();
      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero degli alert:", error);
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
      const items = await this.sp.web.lists
        .getByTitle("Files")
        .items.select(
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

      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero dei files:", error);
      throw error;
    }
  }

  /**
   * Carica un file nella document library 'Files'.
   */
  public async uploadFile(file: File): Promise<void> {
    try {
      // Carica nella root della document library 'Files'
      await this.sp.web
        .getFolderByServerRelativePath("Files")
        .files.addUsingPath(file.name, file, { Overwrite: true });
    } catch (error) {
      console.error("Errore durante il caricamento del file:", error);
      throw error;
    }
  }

  /**
   * Elimina uno o più file dalla document library 'Files' tramite ID item.
   */
  public async deleteFilesById(ids: number[]): Promise<void> {
    try {
      for (const id of ids) {
        await this.sp.web.lists.getByTitle("Files").items.getById(id).delete();
      }
    } catch (error) {
      console.error("Errore durante l'eliminazione dei file:", error);
      throw error;
    }
  }
}
