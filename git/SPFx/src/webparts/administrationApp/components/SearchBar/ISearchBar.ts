import { SearchType } from '../../models/AdministrationAppModels';

export interface ISearchBarprops {
    onSelectSearchType: (searchType: SearchType) => void;
    onSearch: (searchText: string) => void;
    onResetFilters: () => void;
    showTypeSelector: boolean;
    enableRemoveFilters:boolean;
    isMobile:boolean;
    searchType:SearchType;
}

export interface ISearchBarState {
    searchBoxText: string;
}