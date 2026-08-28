export interface ISearchBarprops {
    onSearch: (searchText: string) => void;
    onResetFilters: () => void;
    enableRemoveFilters:boolean;
    isMobile:boolean;
}

export interface ISearchBarState {
    searchBoxText: string;
}