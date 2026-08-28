import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../service/BackendService";

export interface INextEventsprops {
  spService: ISPService;
  bkService: IBackendService;
  context: WebPartContext;
  title: string;
  slidesToShow: number;
  slidesToScroll: number;
  moreLink: string;
  locale: string;
}