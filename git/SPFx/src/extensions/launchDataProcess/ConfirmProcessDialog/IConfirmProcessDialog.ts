import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { IBackendService } from "../../../service/BackendService";
import { ISPService } from "../../../service/Service";

export interface IConfirmprocessDialogprops {
    spService: ISPService;
    context: ListViewCommandSetContext;
    bkService: IBackendService;
    isOpen?: boolean;
    isComplete:boolean;
    onClosePanel: (isComplete:boolean) => void;
}

export interface IConfirmprocessDialogState {
    running: boolean;
    error: boolean;
    open: boolean;
    loading: boolean;
    launched: boolean;
}
