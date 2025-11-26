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
import { MSGraphClientV3 } from '@microsoft/sp-http';

export type Project = Record<string, any>;

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
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
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
      const result = await client.api('/users')
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
        .top(999)
        .get();
      return (result.value || []).map((u: any) => ({
        text: u.displayName,
        secondaryText: u.mail || u.userPrincipalName,
        tertiaryText: u.jobTitle,
        id: u.id,
        data: {
          department: u.department,
          userPrincipalName: u.userPrincipalName
        }
      }));
    } catch (error) {
      console.error("Errore nel recupero utenti da Graph:", error);
      return [];
    }
  }

  /**
   * Recupera un singolo utente per ID
   */
  public async getUserById(userId: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const user = await client.api(`/users/${userId}`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones')
        .get();
      return user;
    } catch (error) {
      console.error("Errore nel recupero utente da Graph:", error);
      return null;
    }
  }

  /**
   * Recupera un utente per email
   */
  public async getUserByEmail(email: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const result = await client.api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
        .get();
      return result.value && result.value.length > 0 ? result.value[0] : null;
    } catch (error) {
      console.error("Errore nel recupero utente per email da Graph:", error);
      return null;
    }
  }

  /**
   * Recupera l'utente corrente
   */
  public async getCurrentUser(): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const user = await client.api('/me')
        .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation')
        .get();
      return user;
    } catch (error) {
      console.error("Errore nel recupero utente corrente da Graph:", error);
      return null;
    }
  }
}

/*===================================================================================
    USER PROFILE SERVICE
====================================================================================*/

export class UserProfileService {
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  /**
   * Recupera la URL dell'immagine profilo di un utente dato l'ID
   * @param userId ID dell'utente da Microsoft Graph
   * @returns URL dell'immagine profilo (blob) o stringa vuota
   */
  public async getUserProfilePicture(userIdOrEmail: string): Promise<string> {
  console.log("🚀 getUserProfilePicture called with:", userIdOrEmail);

  try {
    const client = await this.getGraphClient();

    const guidRegex = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
    let userKey = userIdOrEmail;

    if (!guidRegex.test(userIdOrEmail)) {
      console.log("ℹ️ Input is not a GUID, trying to resolve via Graph...");
      
      const result = await client.api('/users')
        .filter(`mail eq '${userIdOrEmail}' or userPrincipalName eq '${userIdOrEmail}'`)
        .select('id,displayName,mail')
        .get();

      console.log("🔍 Graph lookup result:", result);

      if (result.value && result.value.length > 0) {
        userKey = result.value[0].id;
        console.log("✅ Resolved userKey (GUID) for Graph:", userKey);
      } else {
        console.warn("⚠️ User not found in Graph for:", userIdOrEmail);
        return ""; // oppure throw error se vuoi bloccare
      }
    } else {
      console.log("ℹ️ Input is already a GUID:", userKey);
    }

    const photoBlob = await client.api(`/users/${userKey}/photo/$value`).get();
    const url = window.URL.createObjectURL(photoBlob);
    console.log("✅ Photo URL created:", url);
    return url;

  } catch (error) {
    console.error("❌ Errore nel recupero della foto profilo:", error);
    return "";
  }
}


  /**
   * Recupera la foto profilo dell'utente corrente
   */
  public async getCurrentUserProfilePicture(): Promise<string> {
    try {
      const client = await this.getGraphClient();
      const photoBlob = await client.api('/me/photo/$value').get();
      const url = window.URL.createObjectURL(photoBlob);
      return url;
    } catch (error) {
      console.error("Errore nel recupero della foto profilo corrente:", error);
      return "";
    }
  }

  /**
   * Recupera il profilo completo di un utente
   */
  public async getUserProfile(userId: string): Promise<any> {
    try {
      const client = await this.getGraphClient();
      const profile = await client.api(`/users/${userId}`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,city,country,postalCode,state,streetAddress')
        .get();
      return profile;
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
      const manager = await client.api(`/users/${userId}/manager`)
        .select('id,displayName,mail,userPrincipalName,jobTitle')
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
      const result = await client.api(`/users/${userId}/directReports`)
        .select('id,displayName,mail,userPrincipalName,jobTitle')
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
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
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
      const result = await client.api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
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
    DOCUMENTS SERVICE (con espansione utenti via Graph)
====================================================================================*/

export class DocumentsService {
  private sp: SPFI;
  private context: WebPartContext;
  private graphClient: MSGraphClientV3 | null = null;

  constructor(context: WebPartContext) {
    this.sp = spfi().using(SPFx(context));
    this.context = context;
  }

  private async getGraphClient(): Promise<MSGraphClientV3> {
    if (!this.graphClient) {
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;
      
      const client = await this.getGraphClient();
      const result = await client.api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
        .get();
      
      return result.value && result.value.length > 0 ? result.value[0] : null;
    } catch (error) {
      console.error("Errore nell'arricchimento dati utente:", error);
      return null;
    }
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

      // ...existing code...

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

  public async getIssuePurposeChoices(): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle("Documents")
        .fields.getByInternalNameOrTitle("IssuePurpose")();
      return field.Choices as string[];
    } catch (error) {
      console.error("Errore nel recupero delle opzioni IssuePurpose:", error);
      return [];
    }
  }

  public async getApprovalCodeChoices(): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle("Documents")
        .fields.getByInternalNameOrTitle("ApprovalCode")();
      return field.Choices as string[];
    } catch (error) {
      console.error("Errore nel recupero delle opzioni ApprovalCode:", error);
      return [];
    }
  }

  public async getStatusChoices(): Promise<string[]> {
    try {
      const field = await this.sp.web.lists
        .getByTitle("Documents")
        .fields.getByInternalNameOrTitle("Status")();
      return field.Choices as string[];
    } catch (error) {
      console.error("Errore nel recupero delle opzioni Status:", error);
      return [];
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
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;
      
      const client = await this.getGraphClient();
      const result = await client.api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
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
      this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    }
    return this.graphClient;
  }

  private async enrichUserData(userId: number, email?: string): Promise<any> {
    try {
      if (!email) return null;
      
      const client = await this.getGraphClient();
      const result = await client.api('/users')
        .filter(`mail eq '${email}' or userPrincipalName eq '${email}'`)
        .select('id,displayName,mail,userPrincipalName,jobTitle,department')
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