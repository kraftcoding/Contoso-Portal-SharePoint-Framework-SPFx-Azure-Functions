import { ISPService } from "../../../service/Service";

export interface IRecentDocumentsprops {
  spService: ISPService;
  title: string;
  docsToShow: number;
  moreLink: string;
}