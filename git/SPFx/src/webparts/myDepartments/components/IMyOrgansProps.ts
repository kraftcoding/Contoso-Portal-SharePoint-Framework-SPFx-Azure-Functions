import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../service/BackendService";

export interface ImyDepartmentsprops {
  spService: ISPService;
  context: WebPartContext;
  title: string;
  slidesToShow: number;
  slidesToScroll: number;
  bkService: IBackendService;
  locale: string;
}