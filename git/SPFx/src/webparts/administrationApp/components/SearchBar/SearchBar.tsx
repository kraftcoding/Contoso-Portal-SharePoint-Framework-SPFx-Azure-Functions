import * as React from 'react';
import styles from './SearchBar.module.scss';
import {
  ISearchBarprops,
  ISearchBarState
} from './ISearchBar';
import {
  Button,
  Dropdown,
  Input,
  Option,
  OptionOnSelectData,
  SelectionEvents,
  Tooltip
} from '@fluentui/react-components';
//import { SearchBox } from '@microsoft/mgt-react';
import { SearchType } from '../../models/AdministrationAppModels';
import { SearchRegular, TextGrammarDismissRegular } from '@fluentui/react-icons';
import strings from 'AdministrationAppWebPartStrings';


class SearchBar extends React.Component<ISearchBarprops, ISearchBarState> {

  // private _searchType: string | undefined;
  //private _searchBoxText: string;

  constructor(props: ISearchBarprops) {
    super(props);
    this.state = { searchBoxText: "" };
  }

  private onChangeSearchType(inputText: string): void {
    const { onSelectSearchType } = this.props;

    if (inputText === SearchType.Organ) {
      onSelectSearchType(SearchType.Organ);
    }
    else if (inputText === SearchType.User) {
      onSelectSearchType(SearchType.User);
    }
  }

  private onSearch(inputText: string): void {
    const { onSearch } = this.props;
    onSearch(inputText);
    if (inputText === "") {
      this.setState({ searchBoxText: "" });
    }
  }

  private onResetFilters(): void {
    const { onResetFilters } = this.props;
    onResetFilters();
    this.setState({ searchBoxText: "" });
  }

  public render(): JSX.Element {
    const { showTypeSelector, enableRemoveFilters, isMobile, searchType } = this.props;
    const { searchBoxText } = this.state;
    return (
      <div className={styles.searchBar}>
        {/* Tipo de búsqueda */}
        {showTypeSelector &&
          <Dropdown
            className={styles.typeSelector}
            placeholder={strings.SearchTypePlaceholder}
            value={searchType}
            defaultValue={searchType}
            defaultSelectedOptions={[searchType]}
            onOptionSelect={
              (_event: SelectionEvents, data: OptionOnSelectData): void =>
                this.onChangeSearchType(data.optionValue ? data.optionValue.toString() : "")
            }
          >
            <Option
              key={SearchType.Organ}
              value={SearchType.Organ}
            >
              {strings.OrgansType}
            </Option>
            <Option
              key={SearchType.User}
              value={SearchType.User}
            >
              {strings.UsersType}
            </Option>
          </Dropdown>
        }
        <Input contentBefore={<SearchRegular />} placeholder={strings.SearchPlaceholder} className={styles.searchBox} value={searchBoxText} onChange={(ev, data) => { this.setState({ searchBoxText: data?.value }) }}
          onKeyUp={e => {
            if (e.key === 'Enter') {
              this.onSearch(searchBoxText)
            }
          }}
        />
        {/* Botón de buscar */}
        <Button
          className={styles.searchButton}
          onClick={() => this.onSearch(searchBoxText)}
          appearance='primary'
        >
          {strings.SearchLabel}
        </Button>
        {isMobile ?
          <Button disabled={!enableRemoveFilters} aria-label={strings.CleanSearch} appearance={"secondary"} onClick={() => this.onResetFilters()}>{strings.CleanSearch}</Button>
          :
          <Tooltip
            withArrow
            positioning={'below'}
            content={strings.CleanSearch}
            relationship="label"
          >
            <Button disabled={!enableRemoveFilters} appearance={"secondary"} icon={<TextGrammarDismissRegular />} onClick={() => this.onResetFilters()} />
          </Tooltip>
        }
      </div>
    );
  }

}

export default SearchBar;