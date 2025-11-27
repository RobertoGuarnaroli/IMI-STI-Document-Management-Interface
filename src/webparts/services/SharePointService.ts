import { spfi, SPFx, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPersonaProps } from "@fluentui/react/lib/Persona";
import { MSGraphClientV3 } from "@microsoft/sp-http";

export type Project = Record<string, any>;

export class ChoiceFieldService {
  private sp: SPFI;
  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
  }
    /**
   * Recupera le opzioni di scelta per un campo di una lista
   */
  public async getFieldChoices(listName: string, fieldName: string): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle(listName)
        .fields.getByInternalNameOrTitle(fieldName)();
      return field.Choices as string[];
    } catch (error) {
      console.error(`Errore nel recupero delle opzioni per ${fieldName} in ${listName}:`, error);
      return [];
    }
  }
}

// ====================
// 📌 USERS SERVICE
// ====================
export class UsersService {
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient("3");
    }
    return this.graphClient;
  }

  /**
   * Recupera gli utenti da Microsoft Graph (solo utenti, non gruppi)
   * in formato IPersonaProps per il People Picker.
   */
  public async getUsers(): Promise<IPersonaProps[]> {
    try {
      const client = await this.getGraphClient();
      // Recupera i primi 999 utenti (modifica $top se necessario)
      const result = await client
        .api("/users")
        .select("id,displayName,mail,userPrincipalName,jobTitle,department")
        .top(999)
        .get();
      return (result.value || []).map((u: any) => ({
        text: u.displayName,
        secondaryText: u.mail || u.userPrincipalName,
        tertiaryText: u.jobTitle,
        id: u.id,
        data: {
          department: u.department,
          userPrincipalName: u.userPrincipalName,
        },
      }));
    } catch (error) {
      console.error("Errore nel recupero utenti da Graph:", error);
      return [];
    }
  }


  /**
   * Recupera la URL dell'immagine profilo di un utente dato la mail o UPN
   * @param email email o userPrincipalName
   * @returns URL dell'immagine profilo (blob) o stringa vuota
   */
  public async getUserProfilePictureByEmail(email: string): Promise<string> {
    try {
      const client = await this.getGraphClient();
      // Trova l'utente per email o UPN
      const result = await client
        .api("/users")
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select("id")
        .get();
      if (!result.value || result.value.length === 0) return "";
      const userId = result.value[0].id;
      const photoBlob = await client.api(`/users/${userId}/photo/$value`).get();
      const url = window.URL.createObjectURL(photoBlob);
      return url;
    } catch (error) {
      return "";
    }
  }

  /**
   * Recupera il profilo completo di un utente dato la mail o UPN
   */
  public async getUserProfileByEmail(email: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const result = await client
        .api("/users")
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select(
          "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,city,country,postalCode,state,streetAddress"
        )
        .get();
      if (!result.value || result.value.length === 0) return null;
      return result.value[0];
    } catch (error) {
      console.error("Errore nel recupero del profilo utente:", error);
      return null;
    }
  }

  /**
   * Recupera il manager di un utente
   */
  public async getUserManager(userId: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const manager = await client
        .api(`/users/${userId}/manager`)
        .select("id,displayName,mail,userPrincipalName,jobTitle")
        .get();
      return manager;
    } catch (error) {
      console.error("Errore nel recupero del manager:", error);
      return null;
    }
  }

  /**
   * Recupera i report diretti di un utente
   */
  public async getUserDirectReports(userId: string): Promise<any[]> {
    try {
      const client = await this.getGraphClient();
      const result = await client
        .api(`/users/${userId}/directReports`)
        .select("id,displayName,mail,userPrincipalName,jobTitle")
        .get();
      return result.value || [];
    } catch (error) {
      console.error("Errore nel recupero dei report diretti:", error);
      return [];
    }
  }
}
/*===================================================================================
    PROJECTS SERVICE (con espansione utenti via Graph)
====================================================================================*/

export class ProjectsService {
  private sp: SPFI;
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient("3");
    }
    return this.graphClient;
  }

  /**
   * Arricchisce i dati degli utenti con informazioni da Graph
   */
  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;

      const client = await this.getGraphClient();
      const result = await client
        .api("/users")
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select("id,displayName,mail,userPrincipalName,jobTitle,department")
        .get();

      return result.value && result.value.length > 0 ? result.value[0] : null;
    } catch (error) {
      console.error("Errore nell'arricchimento dati utente:", error);
      return null;
    }
  }

  public async getProjects(): Promise<Project[]> {
    try {
      const items = await this.sp.web.lists
        .getByTitle("Projects")
        .items.select(
          "*",
          "ProjectManager/Id",
          "ProjectManager/Title",
          "ProjectManager/EMail",
          "Author/Id",
          "Author/Title",
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail"
        )
        .expand("ProjectManager", "Author", "Editor")();

      // ...existing code...

      // Arricchisci i dati con informazioni da Graph
      const enrichedItems = await Promise.all(
        items.map(async (item) => {
          const enrichedItem = { ...item };

          if (item.ProjectManager) {
            const graphData = await this.enrichUserData(
              item.ProjectManager.Id,
              item.ProjectManager.EMail
            );
            enrichedItem.ProjectManagerGraph = graphData;
          }

          return enrichedItem;
        })
      );

      return enrichedItems as Project[];
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

  /**
   * Recupera le opzioni di scelta per un campo di una lista
   */
  public async getFieldChoices(listName: string, fieldName: string): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle(listName)
        .fields.getByInternalNameOrTitle(fieldName)();
      return field.Choices as string[];
    } catch (error) {
      console.error(`Errore nel recupero delle opzioni per ${fieldName} in ${listName}:`, error);
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
    DOCUMENTS SERVICE (con espansione utenti via Graph)
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
          "AssignedTo/EMail",
          "Author/Id",
          "Author/Title",
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail"
        )
        .expand("AssignedTo", "Author", "Editor")();

      return items;
    } catch (error) {
      console.error("Errore nel recupero dei documenti:", error);
      throw error;
    }
  }

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

      public async updateDocument(
      itemId: number,
      document: {
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
      }
    ): Promise<void> {
      try {
        await this.sp.web.lists
          .getByTitle("Documents")
          .items.getById(itemId)
          .update({
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
        console.error("Errore durante l'aggiornamento del documento:", error);
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
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail"
        )
        .expand("ProjectCode", "Author", "Editor")();

      // ...existing code...
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
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail"
        )
        .expand("ProjectCode", "Author", "Editor")();

      // ...existing code...
      return items;
    } catch (error) {
      console.error("Errore nel recupero delle distribution lists:", error);
      throw error;
    }
  }
}

/*===================================================================================
    DOCUMENT HISTORY SERVICE (con espansione utenti via Graph)
====================================================================================*/

export class DocumentHistoryService {
  private sp: SPFI;
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient("3");
    }
    return this.graphClient;
  }

  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;

      const client = await this.getGraphClient();
      const result = await client
        .api("/users")
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select("id,displayName,mail,userPrincipalName,jobTitle,department")
        .get();

      return result.value && result.value.length > 0 ? result.value[0] : null;
    } catch (error) {
      console.error("Errore nell'arricchimento dati utente:", error);
      return null;
    }
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
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail"
        )
        .expand("DocumentId", "PerformedBy", "Author", "Editor")();

      // ...existing code...

      // Arricchisci i dati con informazioni da Graph
      const enrichedItems = await Promise.all(
        items.map(async (item) => {
          const enrichedItem = { ...item };

          if (item.PerformedBy) {
            const graphData = await this.enrichUserData(
              item.PerformedBy.Id,
              item.PerformedBy.EMail
            );
            enrichedItem.PerformedByGraph = graphData;
          }

          return enrichedItem;
        })
      );

      return enrichedItems;
    } catch (error) {
      console.error("Errore nel recupero della cronologia documenti:", error);
      throw error;
    }
  }
}

/*===================================================================================
    ALERTS SERVICE (con espansione utenti via Graph)
====================================================================================*/

export class AlertsService {
  private sp: SPFI;
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient("3");
    }
    return this.graphClient;
  }

  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;

      const client = await this.getGraphClient();
      const result = await client
        .api("/users")
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select("id,displayName,mail,userPrincipalName,jobTitle,department")
        .get();

      return result.value && result.value.length > 0 ? result.value[0] : null;
    } catch (error) {
      console.error("Errore nell'arricchimento dati utente:", error);
      return null;
    }
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

      // Arricchisci i dati con informazioni da Graph
      const enrichedItems = await Promise.all(
        items.map(async (item) => {
          const enrichedItem = { ...item };

          if (item.AssignedTo) {
            const graphData = await this.enrichUserData(
              item.AssignedTo.Id,
              item.AssignedTo.EMail
            );
            enrichedItem.AssignedToGraph = graphData;
          }

          return enrichedItem;
        })
      );

      return enrichedItems;
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
          "Author/EMail",
          "Editor/Id",
          "Editor/Title",
          "Editor/EMail",
          "CheckoutUser/Id",
          "CheckoutUser/Title",
          "CheckoutUser/EMail"
        )
        .expand("Author", "Editor", "CheckoutUser")();

      console.log("Raw SharePoint items:", items);
      return items;
    } catch (error) {
      console.error("Errore nel recupero dei files:", error);
      throw error;
    }
  }

  public async uploadFile(file: File): Promise<void> {
    try {
      await this.sp.web
        .getFolderByServerRelativePath("Files")
        .files.addUsingPath(file.name, file, { Overwrite: true });
    } catch (error) {
      console.error("Errore durante il caricamento del file:", error);
      throw error;
    }
  }

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
