import { Term } from "../../models/ITag";
import { IErrorsForm, IManagedItem, IOrganResult, IPermissionCheckResult, IUserPermissionsCheck, IUserResult, IUsersInOrgan, IUsersOrganRelation, ILicenseInfo, SearchType, componentTypeMode } from "../../webparts/administrationApp/models/AdministrationAppModels";

export interface IDepartmentDetailsprops{
    managedItem: IManagedItem | any;
    onManageItem: (itemId: string, itemType: SearchType, action: string, item?: IOrganResult | IUserResult| IUsersOrganRelation[]) => void;
    organTypeFilterOptions: Term[];
    ministryFilterOptions: Term[];
    organFilterOptions: Term[];
    ccaaFilterOptions: Term[];
    roleFilterOptions:Term[];
    secretariaOptions:Term[];
    legistatureOptions:Term[];
    BusinessAreaOptions:Term[];
    isMobile?:boolean;
    userDictionary: Record<string, IUserResult>;
    roleDictionary: Record<string, string>;
    organDictionary: Record<string, string>;
    organTypeDictionary: Record<string, string>;
    ministryDictionary: Record<string, string>;
    errorUpdating?:boolean;
    onCheckPermissions: (item:IUserPermissionsCheck) => promise<IPermissionCheckResult>;
    licensesInformation?:ILicenseInfo[];
    onUpdateLicense: (upn: string, license: string) => void;
    showMode:componentTypeMode;
    showConfirmOnPanel?:boolean;
}

export interface IDepartmentDetailsState {
    selectedTab: string;
    editSelectedItem: IOrganResult| IUserResult | any;
    newValues:IUsersInOrgan;
    showConfirmation?:boolean;
    showLoadingConfirmation?:boolean;
    formErrors?:IErrorsForm;
    showCheckPermissionsModal:boolean;
    checkPermissionsInfo?:IUserPermissionsCheck;
    currentLicense?:string;
}