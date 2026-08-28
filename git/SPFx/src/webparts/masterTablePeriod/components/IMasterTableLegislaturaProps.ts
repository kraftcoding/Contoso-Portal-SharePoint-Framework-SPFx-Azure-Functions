import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMasterTablePeriodprops {
  title: string;
  context: WebPartContext;
  pageSize?: number;
}
