import * as React from 'react';
import styles from './SearchBar.module.scss';
import {
  ISearchBarprops,
  ISearchBarState
} from './ISearchBar';
import {
  Button,
  Input,
  Tooltip
} from '@fluentui/react-components';
//import { SearchBox } from '@microsoft/mgt-react';
import { SearchRegular, TextGrammarDismissRegular } from '@fluentui/react-icons';
import strings from 'AdministrationAppWebPartStrings';


class SearchBar extends React.Component<ISearchBarprops, ISearchBarState> {

  // private _searchType: string | undefined;
  //private _searchBoxText: string;

  constructor(props: ISearchBarprops) {
    super(props);
    this.state = { searchBoxText: "" };
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
    const { enableRemoveFilters, isMobile } = this.props;
    const { searchBoxText } = this.state;
    return (
      <div className={styles.searchBar}>
        {/* Tipo de búsqueda */}
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