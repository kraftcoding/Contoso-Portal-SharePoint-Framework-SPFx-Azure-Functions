import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../service/BackendService";
import { ISPService } from "../../../service/Service";
import { ILookupEXTERNALField, IManagedItem, IOrganEXTERNALResult, ISearchFilters} from "../models/CertificationAppModels";
import { Term } from "../../../models/ITag";

export interface ICertificationAppprops {
  context: WebPartContext;
  spService: ISPService;
  bkService: IBackendService;
  localLanguage: string;
  componentTitle: string;
  isMobile:boolean;
  itemsPerPage:number;
}

export interface ICertificationAppState {
  searchText: string;
  organTypeFilterOptions: ILookupEXTERNALField[];
  ministryFilterOptions: ILookupEXTERNALField[];
  organFilterOptions: Term[];
  roleFilterOptions:Term[];
  StatusFilterOptions:ILookupEXTERNALField[];
  searchFilters: ISearchFilters;
  enableRemoveFilters: boolean;
  dialogIsOpen: boolean;
  managedItem: IManagedItem|undefined;
  currentPage:number;
  BusinessAreaOptions:Term[];
  isLoading:boolean;
  organBackList:IOrganEXTERNALResult[];
  organBackResultList:IOrganEXTERNALResult[];
  showLoadDataError:boolean;
  loadDataError: string | null;
  isLoadingUpdates:boolean;
  errorUpdating:boolean;
  documentDialog:boolean;
}