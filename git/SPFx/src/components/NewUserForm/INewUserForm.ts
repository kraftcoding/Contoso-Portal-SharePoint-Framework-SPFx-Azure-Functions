import { Term } from "../../models/ITag";
import { IBackendService } from "../../service/BackendService";
import { BodyUserInfo, NewUserErrorsForm } from "../../webparts/administrationApp/models/AdministrationAppModels";

export interface INewUserFormState {
    userInfo?: BodyUserInfo;
    formErrors?:NewUserErrorsForm;
    showConfirmPanel:boolean;
    isLoadingSave:boolean;
    showError?: boolean;
}

export interface INewUserFormprops {
    organFilterOptions: Term [];
    ccaaFilterOptions: Term [];
    roleFilterOptions:Term[];
    onCreateUser: (user:BodyUserInfo) => void;
    onClosePanel: () => void;
    hasErrorNewUser?:boolean;
    testSelectedOrgan?:Term;
    bkService: IBackendService;
    invitationMsg: string;
}
