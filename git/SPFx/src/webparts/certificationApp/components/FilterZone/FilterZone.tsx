import * as React from 'react';
import styles from './FilterZone.module.scss';
import {
  IFilterZoneprops,
  IFilterZoneState
} from './IFilterZone';

import {
  Button,
  Combobox,
  Input,
  Option,
  OptionOnSelectData,
  SelectionEvents
} from '@fluentui/react-components';
import {
  Certification,
  ILookupEXTERNALField,
  ISearchFilters,
  SearchFilterKeys
} from '../../models/CertificationAppModels';
import { CalendarSearchRegular, Dismiss12Regular, SearchRegular } from '@fluentui/react-icons';
import strings from 'CertificationAppWebPartStrings';
import { Term } from '../../../../models/ITag';
import _ from 'lodash';
import format from 'date-fns/format';

class FilterZone extends React.Component<IFilterZoneprops, IFilterZoneState> {

  constructor(props: IFilterZoneprops) {

    super(props);
    this.state = {
      currentSearchFilters: Certification.eCNTyFilters
    };
  }

  public componentDidUpdate(testvprops: Readonly<IFilterZoneprops>): void {
    if (!_.isEqual(testvprops.searchFilters, this.props.searchFilters)) {
      this.setState({ currentSearchFilters: this.props.searchFilters });
    }
  }

  private onChangeComboBoxValue = (key: SearchFilterKeys, value: string[] | undefined): void => {
    const { searchFilters } = this.props;
    const updatedSearchFilters = { ...searchFilters, [key]: value ? value : [] };

    this.onUpdateResults(updatedSearchFilters);
  }

  private onChangeSingleValues = (key: SearchFilterKeys, value: string | undefined): void => {
    const { searchFilters } = this.props;
    const updatedSearchFilters = { ...searchFilters, [key]: value ? value : "" };
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

  private onTagSimpleClick = (key: SearchFilterKeys): void => {
    // remove selected option
    const { searchFilters } = this.props;
    const updatedSearchFilters = { ...searchFilters, [key]: "" };
    this.onUpdateResults(updatedSearchFilters);
  };

  public render(): JSX.Element {
    const { organTypeFilterOptions, BusinessAreaOptions } = this.props;
    return (<>
      <div
        className={styles.filterZone}
      >
        {
          /* Filtros para el listado de órganos */
          this.renderLookupComboBox(SearchFilterKeys.OrganTypeEXTERNAL, strings.OrganTypeFilter, organTypeFilterOptions, strings.OrganTypeFilterPlaceholder)
        }
        {
          this.renderTaxonomyComboBox(SearchFilterKeys.BusinessArea, strings.BusinessAreaFilter, BusinessAreaOptions, strings.BusinessAreaPlaceholder)
        }
        {
          this.renderTextFilter(SearchFilterKeys.CodigoDepartment, strings.CodigoDepartment, strings.CodigoDepartment)
        }
      </div>
      <div
        className={styles.filterZone}
      >
        <div className={styles.dateFiltersContainer}>
          <label className={styles.label}>
            {strings.CreationDate}
          </label>
          {
            this.renderDateFilter(SearchFilterKeys.FechaConstitucionDesde, strings.CreationDate, strings.DesdeFilter)
          }
          {
            this.renderDateFilter(SearchFilterKeys.FechaConstitucionHasta, strings.CreationDate, strings.HastaFilter)
          }
        </div>
        <div className={styles.dateFiltersContainer}>
          <label className={styles.label}>
            {strings.ExpirationDate}
          </label>
          {
            this.renderDateFilter(SearchFilterKeys.FechaExtincionDesde, strings.ExpirationDate, strings.DesdeFilter)
          }
          {
            this.renderDateFilter(SearchFilterKeys.FechaExtincionHasta, strings.ExpirationDate, strings.HastaFilter)
          }
        </div>
        <div className={styles.dateFiltersContainer}>
          <label className={styles.label}>
            {strings.FechaInscripcion}
          </label>
          {
            this.renderDateFilter(SearchFilterKeys.FechaInscripcionDesde, strings.FechaInscripcion, strings.DesdeFilter)
          }
          {
            this.renderDateFilter(SearchFilterKeys.FechaInscripcionHasta, strings.FechaInscripcion, strings.HastaFilter)
          }
        </div>
      </div>
    </>
    );
  }

  private renderTaxonomyComboBox = (key: SearchFilterKeys, label: string, options: Term[], placeholder: string): React.ReactElement => {
    const { searchFilters } = this.props;
    return (
      <div className={styles.dropdownContainer}>
        <label className={styles.label}>
          {label}
        </label>
            <Combobox
              onOptionSelect={(_event: SelectionEvents, data: OptionOnSelectData): void => {
                this.onChangeComboBoxValue(key, data.selectedOptions)
              }}
              selectedOptions={searchFilters && searchFilters[key] && key === SearchFilterKeys.BusinessArea ? searchFilters[key] : []}
              multiselect
              placeholder={placeholder ? placeholder : strings.SelectValue}
              className={styles.comboBox}
            >
              {
                options.map((option: Term): JSX.Element => (
                  <Option
                    key={option.key}
                    value={option.key}
                    className={styles.optionComboBoxSize}
                  >
                    {option.text}
                  </Option>
                ))
              }
            </Combobox>
            {searchFilters && searchFilters[key] && searchFilters[key].length && key === SearchFilterKeys.BusinessArea ? (
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
                      <label className={styles.tagButtonTxt}>{options.find(op => op.key === option)?.text}</label>
                    </Button>
                  </li>
                ))}
              </ul>
            ) : null}
      </div>
    );
  }

  private renderLookupComboBox = (key: SearchFilterKeys, label: string, options: ILookupEXTERNALField[], placeholder: string): React.ReactElement => {
    const { searchFilters } = this.props;
    return (
      <div className={styles.dropdownContainer}>
        <label className={styles.label}>
          {label}
        </label>
            <Combobox
              onOptionSelect={(_event: SelectionEvents, data: OptionOnSelectData): void => {
                this.onChangeComboBoxValue(key, data.selectedOptions)
              }}
              selectedOptions={searchFilters && searchFilters[key] && key === SearchFilterKeys.OrganTypeEXTERNAL ? searchFilters[key] : []}
              multiselect
              placeholder={placeholder ? placeholder : strings.SelectValue}
              className={styles.comboBox}
            >
              {
                options.map((option: ILookupEXTERNALField): JSX.Element => (
                  <Option
                    key={option.Id}
                    value={option.Id.toString()}
                    className={styles.optionComboBoxSize}
                  >
                    {option.descripcion}
                  </Option>
                ))
              }
            </Combobox>
            {searchFilters && searchFilters[key] && searchFilters[key].length && key === SearchFilterKeys.BusinessArea ? (
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
                      title={options.find(op => op.Id === option)?.descripcion}
                      className={styles.tagButton}
                    >
                      <label className={styles.tagButtonTxt}>{options.find(op => op.Id === option)?.descripcion}</label>
                    </Button>
                  </li>
                ))}
              </ul>
            ) : null}
      </div>
    );
  }

  private renderDateFilter = (key: SearchFilterKeys, label: string, placeholder: string): React.ReactElement => {
    const { currentSearchFilters } = this.state;
    const { searchFilters } = this.props;
    return (
      <div className={[styles.dateSelector, key].join(" ")}>
        <Input
          key={key}
          type={currentSearchFilters && currentSearchFilters[key] ? 'date': 'text'}
          value={currentSearchFilters && currentSearchFilters[key] ? currentSearchFilters[key] as string : ""}
          className={styles.comboBox}
          placeholder={placeholder}
          onFocus={(e) => {
            e.target.type = 'date';

            const input = e.target as HTMLInputElement & {
              showPicker: () => void;
            }

            if (input?.showPicker) {
              input.showPicker();
            } else {
              input?.focus();
            }
          }}
          onBlur={(e) => { e.target.type = e.target.value ? 'date' : 'text' }}
          onChange={(ev, data) => {
            this.setState({ currentSearchFilters: { ...currentSearchFilters, [key]: data.value ? data.value : "" } });
          }}
          contentAfter={
            <div style={{ display: 'flex', width: '48px', justifyContent: 'flex-end' }}>
              {/*currentSearchFilters[key] ?
                <Button title={strings.RemoveFilter} aria-label={label} size="small" icon={<DismissRegular />} onClick={(e) => {
                  this.setState({currentSearchFilters: {...currentSearchFilters,[key]:""}});
                  this.onChangeSingleValues(key, "")
                  const inpt = (e.target as HTMLButtonElement)?.closest(`.${key}`)?.querySelector('input');
                  if(inpt)
                    inpt.type = "text";
                }} style={{ marginRight: '4pt' }}>
                </Button>
                :
                ""*/}
              <Button title={strings.ApplyFilter} aria-label={label} size="small" icon={<CalendarSearchRegular />} onClick={(e) => this.onChangeSingleValues(key, currentSearchFilters[key] ? currentSearchFilters[key] as string : "")}>
              </Button>
            </div>
          }
        />
          {searchFilters && searchFilters[key] ? (
            <ul
              className={styles.tagsList}
            >
              <li key={key}>
                <Button
                  size="small"
                  shape="circular"
                  appearance="primary"
                  icon={<Dismiss12Regular />}
                  iconPosition="after"
                  onClick={() => this.onTagSimpleClick(key)}
                  title={format(new Date(searchFilters[key] as string), 'dd-MM-yyyy')}
                  className={styles.tagButton}
                >
                  <label className={styles.tagButtonTxt}>{format(new Date(searchFilters[key] as string), 'dd-MM-yyyy')}</label>    
                </Button>
              </li>
            </ul>
          ) : null}
      </div>
    );
  }

  private renderTextFilter = (key: SearchFilterKeys, label: string, placeholder: string): React.ReactElement => {
    const { currentSearchFilters } = this.state;
    const { searchFilters } = this.props;

    return (
      <div className={styles.dropdownContainer}>
        <label className={styles.label}>
          {label}
        </label>
        <Input
          className={styles.comboBox}
          placeholder={placeholder}
          value={currentSearchFilters && currentSearchFilters[key] ? currentSearchFilters[key] as string : ""}
          onChange={(ev, data) => {
            if (key === SearchFilterKeys.CodigoDepartment) {
              this.setState({ currentSearchFilters: { ...currentSearchFilters, [key]: data.value ? data.value : "" } })
            }
          }}
          contentAfter={
            <div style={{ display: 'flex', width: '48px', justifyContent: 'flex-end' }}>
              {/*currentSearchFilters[key] ?
                <Button title={strings.RemoveFilter} aria-label={label} size="small" icon={<DismissRegular />} onClick={() => 
                  { 
                    this.setState({currentSearchFilters: {...currentSearchFilters,[key]:""}});
                    this.onChangeSingleValues(key, "")
                  }
                }
                style={{ marginRight: '4pt' }}>
                </Button>
                :
                ""
              */}
              <Button title={strings.ApplyFilter} aria-label={label} size="small" icon={<SearchRegular />} onClick={() => this.onChangeSingleValues(key, currentSearchFilters[key] ? currentSearchFilters[key] as string : "")}>
              </Button>
            </div>
          }
        />
        {searchFilters && searchFilters[key] ? (
          <ul
            className={styles.tagsList}
          >
            <li key={key}>
              <Button
                size="small"
                shape="circular"
                appearance="primary"
                icon={<Dismiss12Regular />}
                iconPosition="after"
                onClick={() => this.onTagSimpleClick(key)}
                title={searchFilters[key] as string}
                className={styles.tagButton}
              >
                <label className={styles.tagButtonTxt}>{searchFilters[key] as string}</label>
              </Button>
            </li>
          </ul>
        ) : null}
      </div>
    );
  }
}

export default FilterZone;