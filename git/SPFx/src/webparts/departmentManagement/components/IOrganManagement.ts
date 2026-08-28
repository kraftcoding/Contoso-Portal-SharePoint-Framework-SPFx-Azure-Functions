import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IManagedItem, IOrganResult, IUserResult } from "../../administrationApp/models/AdministrationAppModels";
import { ISPService } from "../../../service/Service";
import { IBackendService } from "../../../service/BackendService";
import { Term } from "../../../models/ITag";

export interface IdepartmentManagementprops {
  context: WebPartContext;
  spService: ISPService;
  bkService: IBackendService;
  localLanguage: string;
  componentTitle: string;
  isMobile:boolean;
}

export interface IdepartmentManagementState {
  organTypeOptions: Term[];
  ministryOptions: Term[];
  ccaaFilterOptions: Term[];
  Departmentptions: Term[];
  roleOptions:Term[];
  managedItem: IManagedItem|undefined;
  secretariaOptions:Term[];
  legistatureOptions:Term[];
  BusinessAreaOptions:Term[];
  isLoading:boolean;
  organBackList:IOrganResult[];
  userBackList:IUserResult[];
  showLoadDataError:boolean;
  loadDataError: string | null;
  isLoadingUpdates:boolean;
  errorUpdating:boolean;
  newUserDialogIsOpen:boolean;
  newUsersFromExcel:boolean;
  errorCreatingUser:boolean;
}
