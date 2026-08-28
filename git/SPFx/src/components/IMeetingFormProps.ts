import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPService } from "./../service/Service";
import { IBackendService } from "./../service/BackendService";
import { IEvent } from "../service/BackendServiceModels/EventModels";

export interface IMeetingFormprops {
    context: WebPartContext;
    spService: ISPService;
    bkService: IBackendService;
    editMode: boolean;
    eventData?: IEvent;
    openForm: () => void;
    dataChanged: (changedEvent: IEvent) => void;
    isEditor: boolean;
    readOnly: boolean;
}
