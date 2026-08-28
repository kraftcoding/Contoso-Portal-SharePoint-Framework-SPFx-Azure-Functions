import { ISPService } from '../../../service/Service';
import { IBackendService } from '../../../service/BackendService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
    ILicenseInfo,
    IManagedItem,
    IOrganAdministration,
    IOrganResult,
    ISearchFilters,
    IUserAdministration,
    IUserResult,
    SearchType
} from '../models/AdministrationAppModels';
import { Term } from '../../../models/ITag';

export interface IAdministrationAppprops {
    context: WebPartContext;
    spService: ISPService;
    bkService: IBackendService;
    localLanguage: string;
    componentTitle: string;
    componentMode: string;
    isMobile:boolean;
    itemsPerPage:number;
    showEXTERNAL: boolean;
}

export interface IAdministrationAppState {
    searchType: SearchType;
    searchText: string;
    organsList: IOrganAdministration[];
    resultsOrgansList:IOrganAdministration[];
    resultsUsersList:IUserAdministration[];
    usersList: IUserAdministration[];
    organTypeFilterOptions: Term[];
    ministryFilterOptions: Term[];
    organFilterOptions: Term[];
    ccaaFilterOptions: Term[];
    roleFilterOptions:Term[];
    searchFilters: ISearchFilters;
    enableRemoveFilters: boolean;
    dialogIsOpen: boolean;
    managedItem: IManagedItem|undefined;
    currentPage:number;
    secretariaOptions:Term[];
    legistatureOptions:Term[];
    BusinessAreaOptions:Term[];
    isLoading:boolean;
    organBackList:IOrganResult[];
    userBackList:IUserResult[];
    organBackResultList:IOrganResult[];
    userBackResultList:IUserResult[];
    showLoadDataError:boolean;
    loadDataError: string | null;
    isLoadingUpdates:boolean;
    errorUpdating:boolean;
    availableLicenses:ILicenseInfo[];
    newUserDialogIsOpen:boolean;
    newUsersFromExcel:boolean;
    errorCreatingUser:boolean;
}