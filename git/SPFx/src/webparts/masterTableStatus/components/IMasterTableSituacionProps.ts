import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMasterTableStatusprops {
  title: string;
  context: WebPartContext;
  pageSize?: number;
}
