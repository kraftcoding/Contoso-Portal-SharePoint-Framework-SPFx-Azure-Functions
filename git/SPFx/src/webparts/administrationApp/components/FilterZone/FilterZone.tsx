import * as React from 'react';
import styles from './FilterZone.module.scss';
import {
  IFilterZoneprops,
  IFilterZoneState
} from './IFilterZone';
import {
  PeoplePicker
} from '@microsoft/mgt-react/dist/es6/spfx';
import { MgtPeoplePicker, PersonType, UserType } from '@microsoft/mgt-spfx';
import {
  Button,
  Combobox,
  Label,
  Option,
  OptionOnSelectData,
  SelectionEvents
} from '@fluentui/react-components';
import {
  SearchType,
  ISearchFilters,
  SearchFilterKeys,
  LicensesTypes,
  LicenseNames
} from '../../models/AdministrationAppModels';
import { Dismiss12Regular } from '@fluentui/react-icons';
import strings from 'AdministrationAppWebPartStrings';
import { Term } from '../../../../models/ITag';

class FilterZone extends React.Component<IFilterZoneprops, IFilterZoneState> {

  protected pplPicker = React.createRef<MgtPeoplePicker>();

  constructor(props: IFilterZoneprops) {
    super(props);
    this.state = {};
  }
  /*
    public componentDidMount(): void {
      console.log(this.props.searchType);
    }
  */
  public componentDidUpdate(testvprops: Readonly<IFilterZoneprops>): void {
    const { searchFilters } = this.props;
    if (testvprops.searchFilters[SearchFilterKeys.User] !== searchFilters[SearchFilterKeys.User]) {
      if (!searchFilters[SearchFilterKeys.User]) {
        const n = this.pplPicker.current;
        if (n) {
          n.selectedPeople = [];
          const shad = n.shadowRoot;
          const p = shad?.getElementById("people-picker-input");
          const input = p?.shadowRoot?.getElementById("control");
          input?.removeAttribute("disabled");
          //n.focus();
          // n.disabled = false;
        }
      }
    }
  }

  private onChangeDropdownValue = (key: SearchFilterKeys, value: string | undefined): void => {
    const { searchFilters } = this.props;
    const updatedSearchFilters = { ...searchFilters, [key]: value ? value : "" };

    this.setState({ searchFilters: updatedSearchFilters });
    this.onUpdateResults(updatedSearchFilters);
  }

  private onChangeComboBoxValue = (key: SearchFilterKeys, value: string[] | undefined): void => {
    const { searchFilters } = this.props;
    const updatedSearchFilters = { ...searchFilters, [key]: value ? value : [] };

    this.setState({ searchFilters: updatedSearchFilters });
    this.onUpdateResults(updatedSearchFilters);
  }

  private onUpdateResults = (filters: ISearchFilters): void => {
    this.props.onApplyFilters(filters);
  }

  private onTagClick = (key: SearchFilterKeys, option: string, index: number): void => {
    // remove selected option
    const { searchFilters } = this.props;
    const removedResultArray = [...searchFilters[key]];
    const updatedSearchFilters = { ...searchFilters, [key]: removedResultArray.filter(val => val !== option) };
    this.onUpdateResults(updatedSearchFilters);
  };

  public render(): JSX.Element {
    const { organTypeFilterOptions, organFilterOptions, ccaaFilterOptions, roleFilterOptions, searchType, showEXTERNAL, BusinessAreaOptions } = this.props;
    return (
      <div
        className={styles.filterZone}
        style={{ display: searchType === undefined ? 'none' : '' }}
      >
        {
          /* Filtros para el listado de órganos */
          (searchType === SearchType.Organ) &&
          <>
            {this.renderTaxonomyComboBox(SearchFilterKeys.OrganType, strings.OrganTypeFilter, organTypeFilterOptions, strings.OrganTypeFilterPlaceholder)}
            {!showEXTERNAL
              ? this.renderUserSelector(SearchFilterKeys.User, strings.UserFilter)
              :
              this.renderTaxonomyComboBox(SearchFilterKeys.BusinessArea, strings.BusinessAreaFilter, BusinessAreaOptions, strings.BusinessAreaPlaceholder)
            }
          </>
        }
        {
          /* Filtros para el listado de usuarios */
          (searchType === SearchType.User) &&
          <>
            {this.renderTaxonomyComboBox(SearchFilterKeys.Organ, strings.OrganFilter, organFilterOptions, strings.OrganFilterPlaceholder)}
            {this.renderTaxonomyComboBox(SearchFilterKeys.AutonomousCommunity, strings.CCAAFilter, ccaaFilterOptions, strings.CCAAFilterPlaceholder)}
            {this.renderTaxonomyComboBox(SearchFilterKeys.Rol, strings.RoleFilter, roleFilterOptions, strings.RoleFilterPlaceholder)}
            {this.renderComboBox(SearchFilterKeys.License, strings.LicenseFilter, strings.LicenseFilterPlaceholder)}
          </>
        }
      </div>
    );
  }

  private renderTaxonomyComboBox = (key: SearchFilterKeys, label: string, options: Term[], placeholder: string): React.ReactElement => {
    const { searchFilters } = this.props;
    return (
      <div className={styles.dropdownContainer}>
        <label className={styles.label}>
          {label}
        </label>
        {(key === SearchFilterKeys.Ministry) || (key === SearchFilterKeys.BusinessArea) || (key === SearchFilterKeys.Organ) || (key === SearchFilterKeys.AutonomousCommunity) || (key === SearchFilterKeys.Rol) || (key === SearchFilterKeys.OrganType) ?
          <>
            <Combobox
              onOptionSelect={(_event: SelectionEvents, data: OptionOnSelectData): void => {
                this.onChangeComboBoxValue(key, data.selectedOptions)
              }}
              selectedOptions={searchFilters && searchFilters[key] ? searchFilters[key] : []}
              multiselect
              placeholder={placeholder ? placeholder : strings.SelectValue}
              className={styles.comboBox}
            >
              {key !== SearchFilterKeys.AutonomousCommunity ?
                options.map((option: Term): JSX.Element => (
                  <Option
                    key={option.key}
                    value={option.key}
                    className={styles.optionComboBoxSize}
                  >
                    {option.text}
                  </Option>
                ))
                :
                options.map((option: Term): JSX.Element => (
                  <Option
                    key={option.text}
                    value={option.text}
                    className={styles.optionComboBoxSize}
                  >
                    {option.text}
                  </Option>
                ))
              }
            </Combobox>
            {searchFilters && searchFilters[key] && searchFilters[key].length ? (
              <ul
                className={styles.tagsList}
              >
                {searchFilters[key].map((option, i) => (
                  <li key={option}>
                    <Button
                      size="small"
                      shape="circular"
                      appearance="primary"
                      icon={<Dismiss12Regular />}
                      iconPosition="after"
                      onClick={() => this.onTagClick(key, option, i)}
                      title={options.find(op => op.key === option)?.text}
                      className={styles.tagButton}
                    >
                      <label className={styles.tagButtonTxt}>{key === SearchFilterKeys.AutonomousCommunity ? options.find(op => op.text === option)?.text : options.find(op => op.key === option)?.text}</label>
                    </Button>
                  </li>
                ))}
              </ul>
            ) : null}
          </>
          :
          <Combobox
            onOptionSelect={(_event: SelectionEvents, data: OptionOnSelectData): void => {
              this.onChangeDropdownValue(key, data.optionValue)
            }}
            selectedOptions={searchFilters &&  key !==SearchFilterKeys.License && searchFilters[key] ? [searchFilters[key]] : []}
            clearable
            placeholder={placeholder ? placeholder : strings.SelectValue}
            className={styles.comboBox}
          >
            {
              options.map((option: Term): JSX.Element => (
                <Option
                  key={option.key}
                  value={option.key}
                >
                  {option.text}
                </Option>
              ))
            }
          </Combobox>
        }
      </div>
    );
  }


  private renderComboBox = (key: SearchFilterKeys, label: string, placeholder: string): React.ReactElement => {
    const { searchFilters } = this.props;
    const dropdownOptions:{key: string;value: string;}[]=
    [
      {
        key:LicensesTypes.E3, value:LicenseNames[LicensesTypes.E3],
      },
      {
        key:LicensesTypes.E5, value:LicenseNames[LicensesTypes.E5],
      }
    ];
    return (
      <div className={styles.dropdownContainer}>
        <label className={styles.label}>
          {label}
        </label>
          <Combobox
            onOptionSelect={(_event: SelectionEvents, data: OptionOnSelectData): void => {
              this.onChangeComboBoxValue(key, data.selectedOptions)
            }}
            selectedOptions={searchFilters && key ===SearchFilterKeys.License && searchFilters[key] ? searchFilters[key] : []}
            multiselect
            placeholder={placeholder ? placeholder : strings.SelectValue}
            className={styles.comboBox}
          >
            {
              dropdownOptions.map((option): JSX.Element => (
                <Option
                  key={option.key}
                  value={option.key}
                  className={styles.optionComboBoxSize}
                >
                  {option.value}
                </Option>
              ))}
          </Combobox>
          {searchFilters && key ===SearchFilterKeys.License && searchFilters[key] && searchFilters[key].length ? (
            <ul
              className={styles.tagsList}
            >
              {searchFilters[key].map((option, i) => (
                <li key={option}>
                  <Button
                    size="small"
                    shape="circular"
                    appearance="primary"
                    icon={<Dismiss12Regular />}
                    iconPosition="after"
                    onClick={() => this.onTagClick(key, option, i)}
                    title={option}
                    className={styles.tagButton}
                  >
                    <label className={styles.tagButtonTxt}>{option === LicensesTypes.E3 ? LicenseNames[LicensesTypes.E3] : LicenseNames[LicensesTypes.E5]}</label>
                  </Button>
                </li>
              ))}
            </ul>
          ) : null}
      </div>
    );
  }

  private renderUserSelector = (key: SearchFilterKeys, label: string): React.ReactElement => {
    const { searchFilters } = this.props;
    return (
      <div className={styles.dropdownContainer}>
        <Label className={styles.label}>
          {label}
        </Label>
        <PeoplePicker
          key={key}
          type={PersonType.person}
          userType={UserType.user}
          transitiveSearch={true}
          selectionChanged={async (e) => {
            if (e && e.detail && e.detail.length > 0) {
              const upn = e.detail[0] as { userPrincipalName: string };
              const updatedSearchFilters = { ...searchFilters, [key]: upn.userPrincipalName ? upn.userPrincipalName : "" };

              this.setState({ searchFilters: updatedSearchFilters });
              this.onUpdateResults(updatedSearchFilters);
            } else {
              const updatedSearchFilters = { ...searchFilters, [key]: "" };
              this.setState({ searchFilters: updatedSearchFilters });
              this.onUpdateResults(updatedSearchFilters);
            }
          }}
          selectionMode='single'
          //groupId={"85bc6717-b518-4c97-ba11-30056b8bef61"}
          placeholder={strings.UserFilterPlaceholder}
          className={styles.peoplePicker}
          ref={this.pplPicker}
          disabled={false}
        //selectedPeople={searchFilters && searchFilters[key] && this.filteredPerson ? [this.filteredPerson] : undefined}
        />
      </div>
    );
  }

}

export default FilterZone;