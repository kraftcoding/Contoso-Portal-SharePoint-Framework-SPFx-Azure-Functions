import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMasterTableDivisionprops {
  title: string;
  context: WebPartContext;
  pageSize?: number;
}
