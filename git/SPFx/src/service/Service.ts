import { SPFI, spfi, SPFx as spSPFx } from "@pnp/sp";
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import { AadTokenproviderFactory, SPHttpClient, MSGraphClientFactory } from "@microsoft/sp-http";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { Channel, Teams } from "../models/ITeams";
import { AadGroup } from "../models/IAadGroup";
import IQuery, { IQueryOrder } from "../models/IQuery";
import { ITermInfo } from "@pnp/sp/taxonomy";
import { StestcentDocument } from "../models/IRecentDocument";
import "@pnp/graph/users";
import "@pnp/graph/teams";
import "@pnp/graph/groups";
import "@pnp/graph/photos";
import "@pnp/graph/lists";
import "@pnp/graph/onedrive";
import "@pnp/sp/profiles";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/taxonomy";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import { IFolderAddResult, IFolderInfo, folderFromServerRelativePath } from "@pnp/sp/folders";
import { IFileAddResult, IFileInfo } from "@pnp/sp/files";

export interface ISPService {
  getMyTeams(): promise<Teams[]>;
  getMyAadGroups(): promise<AadGroup[]>;
  getTeamsPhoto(teamId: string): promise<Blob>;
  getTeamsInformation(teamId: string): promise<Teams>;
  getTeamsPrimaryChannel(teamId: string): promise<Channel>;
  getMeInfo(): promise<any>;
  getItems<T>(listUrl: string, queryOptions: IQuery | IQueryOrder): promise<T[]>;
  getItemsByOrder<T>(listUrl: string, queryOptions: IQueryOrder): promise<T[]>;
  getTaxonomy(setId: string): promise<ITermInfo[]>;
  getTaxonomyTerm(setId: string, termId: string): promise<ITermInfo>;
  getRecentDocuments(): promise<StestcentDocument[]>;
  getFiles(folderUrl: string, files: IFileInfo[]): promise<IFileInfo[]>;
  getFileInfoById(uniqueId: string): promise<IFileInfo>;
  copyToDocumentSet(itemsToCopy: string[], destinationPath: string): promise<any>;
  getGroupId(groupName: string): promise<string>;
  getpropertiesByRootFolder(): promise<any>;
  getFolderIdsFromPath(relativePath: string, relativeSitePath: string): promise<string[]>;
  getFileContentFromTxt(filePath: string): promise<string>;
  getFoldersFromPath(relativeFolderPath: string): promise<IFolderInfo[]>;
  ensureFolder(relativeFolderPath: string):promise<IFolderAddResult>;
  uploadFile(relativeFolderPath: string, file:any):promise<IFileAddResult>;
}
export default class SPService implements ISPService {
  public static readonly serviceKey: ServiceKey<ISPService> = ServiceKey.create<ISPService>(
    "SPFx:ContosoSPService",
    SPService
  );

  public static readonly MaxTop: number = 5000;
  public static readonly groupId: string = "102a5d61-9d78-42e8-ab11-681331d4d37f"; //  Grupo de términos service-account@contoso.local

  private sp: SPFI;
  private graph: GraphFI;
  private msGraphClientFactory: MSGraphClientFactory;
  private pageContext: PageContext;
  private spHttpClient: SPHttpClient;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this.spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this.msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
      this.pageContext = serviceScope.consume(PageContext.serviceKey);
      const aadTokenproviderFactory = serviceScope.consume(AadTokenproviderFactory.serviceKey);
      this.sp = spfi().using(spSPFx({ pageContext: this.pageContext }));
      this.graph = graphfi().using(graphSPFx({ aadTokenproviderFactory }));
    });
  }

  public async getMyTeams(): promise<Teams[]> {
    return this.graph.me.joinedTeams.select("id, displayName, tenantId")();
  }

  public async getMyAadGroups(): promise<AadGroup[]> {
    return this.graph.me.transitiveMemberOf.select("id, displayName")();
  }

  public async getTeamsPhoto(teamId: string): promise<Blob> {
    const client = await this.msGraphClientFactory.getClient("3");
    return client.api(`/teams/${teamId}/photo/$value`).get();
  }

  public async getTeamsInformation(teamId: string): promise<Teams> {
    return this.graph.teams.getById(teamId).select("webUrl")();
  }

  public async getTeamsPrimaryChannel(teamId: string): promise<Channel> {
    return (await this.graph.teams.getById(teamId).primaryChannel.select("id")())
  }
  public async getMeInfo(): promise<any> {
    return this.sp.profiles.myproperties();
  }

  public async getMe(): promise<any> {
    return this.graph.me();
  }

  public async getItems<T>(listUrl: string, queryOptions: IQuery | IQueryOrder): promise<T[]> {
    return this.sp.web
      .getList(listUrl)
      .items.filter(queryOptions.filter)
      .top(queryOptions.top || SPService.MaxTop)
      .select(queryOptions.viewFields.join(","))
      .expand(queryOptions.expand.join(","))();
  }

  public async getItemsByOrder<T>(listUrl: string, queryOptions: IQueryOrder): promise<T[]> {
    return this.sp.web
      .getList(listUrl)
      .items.filter(queryOptions.filter)
      .orderBy(queryOptions.orderBy, queryOptions.orderByAscending)
      .top(queryOptions.top || SPService.MaxTop)
      .select(queryOptions.viewFields.join(","))
      .expand(queryOptions.expand.join(","))();
  }

  public async getTaxonomy(setId: string): promise<ITermInfo[]> {
    return this.sp.termStore.groups.getById(SPService.groupId).sets.getById(setId).terms();
  }

  public async getTaxonomyTerm(setId: string, termId: string): promise<ITermInfo> {
    return this.sp.termStore.groups.getById(SPService.groupId).sets.getById(setId).getTermById(termId)();
  }

  public async getRecentDocuments(): promise<StestcentDocument[]> {
    const currentUserDrive = await this.graph.me.drive();
    const recentDocuments: StestcentDocument[] = currentUserDrive.id
      ? await this.graph.me.drives
        .getById(currentUserDrive.id)
        .select("name", "remoteItem/webDavUrl")
        .expand("remoteItem")
        .recent()
      : [];
    return recentDocuments;
  }

  public async getFiles(folderUrl: string, files: IFileInfo[] = []): promise<IFileInfo[]> {
    const folder: any = await folderFromServerRelativePath(this.sp.web, folderUrl)
      .select("name,uniqueid")
      .expand("Folders", "Files")();
    folder.Files.map((file: IFileInfo) => files.push(file));
    await promise.all(
      folder.Folders.map((folder: IFolderInfo) => {
        if (folder.Name !== "_hidden") return this.getFiles(folder.ServerRelativeUrl, files);
      })
    );

    return files;
  }

  public async getFileInfoById(uniqueId: string): promise<IFileInfo> {
    return this.sp.web.getFileById(uniqueId).select("UniqueId", "Name", "ServerRelativeUrl", "TimeLastModified")();
  }

  public async copyToDocumentSet(itemsToCopy: string[], destinationPath: string): promise<any> {
    return this.spHttpClient.post("/_api/site/CreateCopyJobs", SPHttpClient.configurations.v1, {
      body: JSON.stringify({
        exportObjectUris: itemsToCopy,
        destinationUri: destinationPath,
        options: {
          IgnoreVersionHistory: true,
          AllowSchemaMismatch: true,
          NameConflictBehavior: 2,
        },
      }),
    });
  }

  public async getGroupId(groupName: string): promise<string> {
    const groups = await this.graph.groups.filter(`displayName eq '${groupName}'`).select("id, displayName")();
    return groups && groups[0] && groups[0].id ? groups[0].id : "";
  }

  public async getpropertiesByRootFolder(): promise<any> {
    const mainLibrary = await this.sp.web.defaultDocumentLibrary();
    return this.sp.web.folders.getByUrl(mainLibrary.EntityTypeName?.replace(/_x0020_/g, ' ')).properties();
  }



  public async getFolderIdsFromPath(relativeFolderPath: string, relativeSitePath: string): promise<any> {
    console.log(relativeSitePath);
    const splittedPath = relativeFolderPath.split("/");

    const folderpromises = [];

    for (let i = splittedPath.length; i > 0; i--) {
      const relativePath = `${relativeSitePath}/${splittedPath.slice(0, i).join("/")}`;
      folderpromises.push(this.sp.web.getFolderByServerRelativePath(relativePath).listItemAllFields());
    }

    const folders = await promise.all(folderpromises);

    return folders?.map(folder => folder.Id);
  }

  public async getFileContentFromTxt(filePath: string): promise<string> {
    const file = this.sp.web.getFileByServerRelativePath(filePath);
    return await file.getText();
  }

  public async getFoldersFromPath(relativeFolderPath: string): promise<IFolderInfo[]> {
    const baseFolder = this.sp.web.getFolderByServerRelativePath(relativeFolderPath);
      // Obtén las subcarpetas (primer nivel)
    return await baseFolder.folders();
  }

  public async ensureFolder(relativeFolderPath: string):promise<IFolderAddResult>{
    return await this.sp.web.folders.addUsingPath(relativeFolderPath);
  }
  
  public async uploadFile(relativeFolderPath: string, file: File): promise<IFileAddResult> {
    console.log(file);
    const folder= folderFromServerRelativePath(this.sp.web, relativeFolderPath);
    //const result = await folder.files.add(file.name, file, true); // true = sobrescribir si existe
    let result: IFileAddResult;
    if (file.size <= 10485760) {
      // small upload
      result = await folder.files.addUsingPath(file.name, file, { Overwrite: true });
    } else {
      // large upload
      result = await folder.files.addChunked(file.name, file);
    }
    return result;
  }

}
