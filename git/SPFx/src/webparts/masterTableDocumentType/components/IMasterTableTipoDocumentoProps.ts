import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMasterTableTipoDocumentprops {
  title: string;
  context: WebPartContext;
  pageSize?: number;
}
