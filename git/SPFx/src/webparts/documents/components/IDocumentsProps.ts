import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentsprops {
  spService: ISPService;
  context: WebPartContext;
  title: string;
  docsToShow: number;
  moreLink: string;
}