import { Term } from '../../../../models/ITag';
import { ILookupEXTERNALField, ISearchFilters } from '../../models/CertificationAppModels';

export interface IFilterZoneprops {
    onApplyFilters: (searchType: ISearchFilters) => void;
    organTypeFilterOptions: ILookupEXTERNALField[];
    BusinessAreaOptions: Term[];
    searchText: string;
    searchFilters: ISearchFilters;
}

export interface IFilterZoneState {
    currentSearchFilters: ISearchFilters;
}