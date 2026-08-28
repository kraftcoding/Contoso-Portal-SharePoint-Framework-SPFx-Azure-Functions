import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../service/BackendService";
import { ISPService } from "../../../service/Service";

export interface IDelegateAdministrationBannerprops {
  context: WebPartContext;
  bkService: IBackendService;
  spService:ISPService;
  componentTitle: string;
}

export interface IDelegateAdministrationBannerState {
  showBanner:boolean;
}

