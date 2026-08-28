import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Term } from "../../models/ITag";
import { IBackendService } from "../../service/BackendService";
import { NewUserFromExcel } from "../../webparts/administrationApp/models/AdministrationAppModels";

export interface INewUsersFromExcelprops {
    isMobile:boolean;
    onRealoadComponent: () => void;
    onClosePanel: () => void;
    organFilterOptions: Term [];
    bkService: IBackendService;
    context: WebPartContext;
    testSelectedOrgan?:Term;
    invitationMsg:string;
}
export interface INewUsersFromExcelState {
    newUsers: NewUserFromExcel[];
    showConfirmPanel:boolean;
    selectedOrgan?: string;
    correctRequest: string[];
    failRequest: string[];
    isprocessing: boolean;
    loading:boolean;
    showErrorUsersNotprocessed?:boolean;
    invitationMessage:string;
    errors: INewUsersErrors[];
}

export interface INewUsersErrors {
    Nombre?:string;
    Apellidos?:string;
    Email?: string;
    Cargo?: string;
    ComunidadAutonoma?: string;
    Rol?: string;
    Division_GobiernoAutonomico?:string;
    Unidad_Consejeria?:string;
    TelefonoMovil?:string;
    TelefonoEmtestsa?:string;
    CorreoElectronicoSecundario?:string;
    TipoEmpleado?:string;
    DestinatarioInvitacionCC?:string;
}
