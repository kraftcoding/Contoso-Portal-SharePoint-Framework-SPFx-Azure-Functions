import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPService } from "../../../../service/Service";
import { IBackendService } from "../../../../service/BackendService";

export interface ICallsprops {
    context: WebPartContext;
    spService: ISPService;
    bkService: IBackendService;
}
