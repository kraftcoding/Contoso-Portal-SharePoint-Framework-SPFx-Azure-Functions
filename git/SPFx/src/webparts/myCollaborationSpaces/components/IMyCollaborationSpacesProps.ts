import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyCollaborationSpacesprops {
  spService: ISPService;
  title: string;
  numberOfResults: number;
  showDescription: boolean;
  moreLink: string;
  context: WebPartContext;
}
