import { ISPService } from '../../../service/Service';
import { IBackendService } from '../../../service/BackendService';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Term } from '../../../models/ITag';
import { ILookupMinistryField, IRelationBusinessArea, IRelationsFilters } from './BusinessAreaRelationsModels';
export interface IBusinessAreaRelationsprops {
  context: WebPartContext;
  spService: ISPService;
  bkService: IBackendService;
  localLanguage: string;
  title: string;
  isMobile:boolean;
}

export interface IBusinessAreaRelationsState {
  allRelations:IRelationBusinessArea[];
  filteredRelations:IRelationBusinessArea[];
  filters:IRelationsFilters;
  ministryFilterOptions: ILookupMinistryField[];
  BusinessAreaFilterOptions:Term[];
  isLoading:boolean;
  showDialog:boolean;
  hasError:boolean;
  itemsWithError:IRelationBusinessArea[];
  isLoadingConfirm:boolean;
  hasErrorUpdating:boolean;
  isNewItemMode:boolean;
  newItem:IRelationBusinessArea|undefined;
}
