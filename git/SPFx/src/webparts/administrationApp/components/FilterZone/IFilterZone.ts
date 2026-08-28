import { Term } from '../../../../models/ITag';
import { SearchType, ISearchFilters } from '../../models/AdministrationAppModels';

export interface IFilterZoneprops {
    searchType: SearchType;
    onApplyFilters: (searchType: ISearchFilters) => void;
    organTypeFilterOptions: Term[];
    BusinessAreaOptions: Term[];
    organFilterOptions: Term [];
    ccaaFilterOptions: Term [];
    roleFilterOptions:Term[];
    searchText: string;
    searchFilters: ISearchFilters;
    showEXTERNAL?:boolean;
}

export interface IFilterZoneState {
}