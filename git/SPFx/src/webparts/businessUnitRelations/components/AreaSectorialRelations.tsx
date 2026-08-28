import * as React from 'react';
import styles from './BusinessAreaRelations.module.scss';
import type { IBusinessAreaRelationsprops, IBusinessAreaRelationsState } from './IBusinessAreaRelations';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { ILookupMinistryField, IRelationBusinessArea, IRelationsFilters, RelationAreaFields, RelationsAreaMinistry, RelationsFilterKeys, RelationsTaxonomyIds } from './BusinessAreaRelationsModels';
import { Term } from '../../../models/ITag';
import { Button, Combobox, createTableColumn, DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell, DataGridRow, Dialog, DialogActions, DialogBody, DialogContent, DialogSurface, DialogTitle, DialogTrigger, Input, Label, Option, Spinner, Switch, TableCell, TableCellLayout, TableColumnDefinition, TableHeader, TableHeaderCell, TableRow, Tooltip } from '@fluentui/react-components';
import format from 'date-fns/format';
import { AddRegular, ErrorCircle20Filled } from '@fluentui/react-icons';
import strings from 'BusinessAreaRelationsWebPartStrings';
import { EXTERNALList } from '../../certificationApp/models/CertificationAppModels';

export default class BusinessAreaRelations extends React.Component<IBusinessAreaRelationsprops, IBusinessAreaRelationsState> {
  private _BusinessAreaDictionary: Record<string, string>;
  private _ministryDictionary: Record<string, string>;
  private _originalItems: IRelationBusinessArea[];
  private _toUpdateItems: IRelationBusinessArea[];

  private areaColumns: TableColumnDefinition<IRelationBusinessArea>[] =
    this.props.isMobile ?
      [
        createTableColumn<IRelationBusinessArea>({
          columnId: "BusinessArea",
          renderHeaderCell: () => {
            return strings.BusinessAreaLabel;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout className={"styles.dataItemTitle"}>
                <Label>{item.BusinessArea && this._BusinessAreaDictionary[item.BusinessArea]}</Label>
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IRelationBusinessArea>({
          columnId: "Division",
          renderHeaderCell: () => {
            return strings.DivisionLabel;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout className={"styles.dataItemTitle"}>
                {item.Division && this._ministryDictionary[item.Division]}
              </TableCellLayout>
            )
          },
        })
      ] :
      [
        createTableColumn<IRelationBusinessArea>({
          columnId: "BusinessArea",
          renderHeaderCell: () => {
            return strings.BusinessAreaLabel;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout title={item.BusinessArea && this._BusinessAreaDictionary[item.BusinessArea]}>
                {item.BusinessArea && this._BusinessAreaDictionary[item.BusinessArea]}
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IRelationBusinessArea>({
          columnId: "StartDate",
          renderHeaderCell: () => {
            return strings.StartDateLabel;
          },
          renderCell: (item) => {
            /*return (
              <TableCellLayout
              >
                {item.StartDate}
              </TableCellLayout>
            );*/
            return (
              <Input
                key='StartDate'
                value={item?.StartDate && format(new Date(item.StartDate), 'yyyy-MM-dd')}
                type='date'
                onChange={async (data): promise<void> => {
                  item[RelationAreaFields.Id] && this.onUpdateValue(RelationAreaFields.StartDate, item[RelationAreaFields.Id], data.target.value ? new Date(data.target.value).toISOString() : "")
                }}
                onKeyDown={(e): void => e.preventDefault()}
                className={styles.dateSelector}
                title={item?.StartDate && format(new Date(item.StartDate), 'dd-MM-yyyy')}
              />
            )
          },
        }),
        createTableColumn<IRelationBusinessArea>({
          columnId: "EndDate",
          renderHeaderCell: () => {
            return strings.EndDateLabel;
          },
          renderCell: (item) => {
            /*return (
              <TableCellLayout
              >
                {item.EndDate}
              </TableCellLayout>
            );*/
            return (
              <Input
                key='EndDate'
                value={item?.EndDate && format(new Date(item.EndDate), 'yyyy-MM-dd')}
                type='date'
                onChange={async (data): promise<void> => {
                  item[RelationAreaFields.Id] && this.onUpdateValue(RelationAreaFields.EndDate, item[RelationAreaFields.Id], data.target.value ? new Date(data.target.value).toISOString() : "")
                }}
                onKeyDown={(e): void => e.preventDefault()}
                className={styles.dateSelector}
                title={item?.EndDate && format(new Date(item.EndDate), 'dd-MM-yyyy')}
                min={item?.StartDate && format(new Date(item.StartDate), 'yyyy-MM-dd')}
              />
            )
          },
        }),
        createTableColumn<IRelationBusinessArea>({
          columnId: "Division",
          renderHeaderCell: () => {
            return strings.DivisionLabel;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout
              >
                <Combobox
                  onOptionSelect={(_event, data): void => {
                    data.optionValue && this.onUpdateValue(RelationAreaFields.Division, item[RelationAreaFields.Id], data.optionValue)
                    //this.onChangeComboBoxValue(RelationsFilterKeys.Division, data.optionValue)
                  }}
                  selectedOptions={item && item.Division ? [item.Division] : []}
                  placeholder={strings.DivisionPlaceholder}
                  className={styles.ministryDropdown}
                  value={item && item.Division && this._ministryDictionary[item.Division]}
                  title={item && item.Division && this._ministryDictionary[item.Division]}
                >
                  {
                    this.state.ministryFilterOptions.map((option: ILookupMinistryField): JSX.Element => (
                      <Option
                        key={option.Id}
                        value={option.Id.toString()}
                      >
                        {option.Title}
                      </Option>
                    ))
                  }
                </Combobox>
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IRelationBusinessArea>({
          columnId: "openItem",
          renderHeaderCell: () => {
            return "";
          },
          renderCell: (item) => {
            const isVisible = this.state.itemsWithError && this.state.itemsWithError.some(el => el[RelationAreaFields.Id] === item[RelationAreaFields.Id]);
            return (
              <Tooltip
                withArrow
                positioning={'below'}
                content={strings.ReviewErrorLabel}
                relationship="label"
              >
                <ErrorCircle20Filled className={styles.errorRow} style={{ visibility: isVisible ? "visible" : "hidden" }} />
              </Tooltip>
            )
          },
        })
      ]
    ;
  constructor(props: IBusinessAreaRelationsprops) {
    super(props);
    this.state = {
      allRelations: [],
      filteredRelations: [],
      filters: RelationsAreaMinistry.eCNTyFilters,
      ministryFilterOptions: [],
      BusinessAreaFilterOptions: [],
      isLoading: true,
      showDialog: false,
      hasError: false,
      itemsWithError: [],
      isLoadingConfirm: false,
      hasErrorUpdating: false,
      isNewItemMode: false,
      newItem: undefined,
    }
  }

  async componentDidMount(): promise<void> {
    await this.getTaxonomies();
    void this.loadData();
  }

  private loadData(isReload?: boolean): void {
    if (isReload) {
      this._toUpdateItems = [];
      this.setState({
        allRelations: [],
        filteredRelations: [],
        filters: RelationsAreaMinistry.eCNTyFilters,
        isLoading: true,
        showDialog: false,
        hasError: false,
        itemsWithError: [],
        isLoadingConfirm: false,
        hasErrorUpdating: false
      })
    }
    this.props.bkService.getAreaRelationsInfo('Contoso').then(relations => {
      const orderedRelations = relations.sort((a, b) => {
        const areaComparison = this._BusinessAreaDictionary[a.BusinessArea].localeCompare(this._BusinessAreaDictionary[b.BusinessArea]);
        if (areaComparison !== 0) {
          return areaComparison;
        }
        return new Date(b.StartDate).getTime() - new Date(a.StartDate).getTime();
      });

      const currentRelations = this.getCurrent(orderedRelations);
      console.log(orderedRelations);
      console.log(currentRelations);
      this._originalItems = orderedRelations;
      this.setState({
        allRelations: orderedRelations,
        filteredRelations: currentRelations,
        isLoading: false
      })
    }).catch(ex => {
      this.setState({ hasError: true });
      console.log(ex)
    })
  }

  private getCurrent(items: IRelationBusinessArea[]): IRelationBusinessArea[] {
    const filteredItems = Object.values(
      items.reduce((acc, item) => {
        const current = acc[item.BusinessArea];
        if (!current || (new Date(item.StartDate) > new Date(current.StartDate))) {
          acc[item.BusinessArea] = item;
        }
        return acc;
      }, {} as { [key: string]: IRelationBusinessArea })
    );
    return filteredItems
  }

  private async getTaxonomies(): promise<void> {
    const { spService } = this.props;

    const [ministryValues, areaSecTerms]: [ILookupMinistryField[], ITermInfo[]] = await promise.all([
      spService.getItemsByOrder<ILookupMinistryField>(`${this.props.context.pageContext.site.serverRelativeUrl}${EXTERNALList.ListDivisionUrl}`, EXTERNALList.getAll()),
      spService.getTaxonomy(RelationsTaxonomyIds.BusinessArea),
    ]);
    const BusinessAreaTermInfo: Term[] = areaSecTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));

    this._ministryDictionary = ministryValues.reduce((acc, item) => {
      acc[item.Id] = item.Title;
      return acc;
    }, {} as Record<string, string>);

    this._BusinessAreaDictionary = BusinessAreaTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);
    this.setState({
      ministryFilterOptions: ministryValues,
      BusinessAreaFilterOptions: BusinessAreaTermInfo
    });
  }

  private onChangeFilterComboBoxValue = (key: RelationsFilterKeys, value: string | undefined): void => {
    const { filters } = this.state;
    let updatedSearchFilters = { ...filters, [key]: value ? [value] : [] };
    if (key === RelationsFilterKeys.BusinessArea) {
      updatedSearchFilters = { ...updatedSearchFilters, Division: [] }
    }
    this.setState({ filters: updatedSearchFilters });
    this.onUpdateResults(updatedSearchFilters);
  }

  private onChangeFilterBooleanValue = (key: RelationsFilterKeys, value: boolean): void => {
    const { filters } = this.state;
    const updatedSearchFilters = { ...filters, [key]: value };
    this.setState({ filters: updatedSearchFilters });
    this.onUpdateResults(updatedSearchFilters);
  }

  private onUpdateResults = (filters: IRelationsFilters): void => {
    const { Actual, BusinessArea, Division } = filters;
    let results = [...this.state.allRelations];
    if (Actual) {
      results = this.getCurrent(results);
    }
    if (BusinessArea && BusinessArea.length > 0) {
      results = results.filter(item =>
        BusinessArea[0] === item.BusinessArea
      );
    }
    if (Division && Division.length > 0) {
      results = results.filter(item =>
        Division[0].toString() === item.Division.toString()
      );
    }
    this.setState({ filteredRelations: results });
  }

  private onUpdateValue = (key: RelationAreaFields, id: string | undefined, value: string): void => {
    const allItems = [...this.state.allRelations];
    const filteredItems = [...this.state.filteredRelations];
    if (id) {
      const selectedItem = allItems.find(item => item[RelationAreaFields.Id] === id);
      console.log(selectedItem);
      console.log(value);
      if (selectedItem) {
        let updatedItem = { ...selectedItem, [key]: value }
        if (key === RelationAreaFields.StartDate) {
          if (new Date(value) > new Date(updatedItem[RelationAreaFields.EndDate])) {
            updatedItem = { ...updatedItem, [RelationAreaFields.EndDate]: value }
          }
        }
        if (key === RelationAreaFields.EndDate) {
          if (new Date(value) < new Date(updatedItem[RelationAreaFields.StartDate])) {
            updatedItem = { ...updatedItem, [RelationAreaFields.StartDate]: value }
          }
        }

        const hasError = this.hasOverlapValidDates(updatedItem);
        const newAllItems = allItems.map(item => item[RelationAreaFields.Id] === id ? { ...updatedItem, HasError: hasError } : item);
        const newFilteredItems = filteredItems.map(item => item[RelationAreaFields.Id] === id ? { ...updatedItem, HasError: hasError } : item);
        console.log(updatedItem);
        this.setState({ allRelations: [...newAllItems], filteredRelations: [...newFilteredItems] })
      }
    } else {
      const { newItem } = this.state;
      if (newItem) {
        let updatedItem = { ...newItem, [key]: value };

        let hasError = false;
        /*
        if (key === RelationAreaFields.StartDate && updatedItem[RelationAreaFields.EndDate]) {
          if (new Date(value) > new Date(updatedItem[RelationAreaFields.EndDate])) {
            updatedItem = { ...updatedItem, [RelationAreaFields.EndDate]: value }
          }
        }
        if (key === RelationAreaFields.EndDate && updatedItem[RelationAreaFields.StartDate]) {
          if (new Date(value) < new Date(updatedItem[RelationAreaFields.StartDate])) {
            updatedItem = { ...updatedItem, [RelationAreaFields.StartDate]: value }
          }
        }
          */
        if (updatedItem && updatedItem[RelationAreaFields.EndDate] && updatedItem[RelationAreaFields.StartDate] && new Date(updatedItem[RelationAreaFields.StartDate]) > new Date(updatedItem[RelationAreaFields.EndDate])) {
          hasError = true
        }

        if (!hasError) {
          hasError = updatedItem[RelationAreaFields.BusinessArea] && updatedItem[RelationAreaFields.StartDate] && updatedItem[RelationAreaFields.EndDate] && this.hasOverlapValidDates(updatedItem) ? true : false;
        }
        console.log(updatedItem);
        this.setState({ newItem: { ...updatedItem, HasError: hasError ? true : false } });

      } else {
        const newItem: IRelationBusinessArea = {
          [RelationAreaFields.BusinessArea]: '',
          [RelationAreaFields.Division]: '',
          [RelationAreaFields.StartDate]: '',
          [RelationAreaFields.EndDate]: ''
        }
        let updatedItem = { ...newItem, [key]: value };
        console.log(updatedItem);
        this.setState({ newItem: updatedItem });
      }
    }
  }

  private hasOverlapValidDates = (item: IRelationBusinessArea): boolean => {
    const { allRelations } = this.state;
    let hasError = false;

    const { BusinessArea, StartDate, EndDate } = item;
    const startDate = new Date(StartDate);
    const endDate = new Date(EndDate);
    hasError = allRelations.some(otherItem => {
      if (otherItem.BusinessArea !== BusinessArea || otherItem.Id === item.Id) {
        return false;
      }
      const otherStartDate = new Date(otherItem.StartDate);
      const otherEndDate = new Date(otherItem.EndDate);

      return (startDate <= otherEndDate && endDate >= otherStartDate);
    });
    return hasError;
  }

  private onGetUpdatedItems = (): IRelationBusinessArea[] => {
    const { allRelations } = this.state;

    const changes = allRelations.filter(item1 => {
      const item2 = this._originalItems.find(item => item.Id === item1.Id);
      return item2 && (
        item1.BusinessArea !== item2.BusinessArea ||
        item1.Division !== item2.Division ||
        item1.StartDate !== item2.StartDate ||
        item1.EndDate !== item2.EndDate
      );
    });
    return changes;
  }

  private onClickDiscard = (): void => {
    this._toUpdateItems = [];
    this.setState({ filteredRelations: this.getCurrent(this._originalItems), allRelations: this._originalItems, itemsWithError: [], filters: RelationsAreaMinistry.eCNTyFilters, });
  }

  private onClickOpenDialog = (): void => {
    const diffItems = this.onGetUpdatedItems();
    console.log(diffItems);
    if (diffItems && diffItems.length > 0) {
      this._toUpdateItems = diffItems;
      const itemsWithError: IRelationBusinessArea[] = [];
      diffItems.forEach(item => {
        if (this.hasOverlapValidDates(item)) {
          itemsWithError.push(item);
        }
      })
      this.setState({ showDialog: true, itemsWithError: itemsWithError, isNewItemMode: false })
    }
  }

  private onClickSaveConfirmation = (): void => {
    this.setState({ isLoadingConfirm: true });

    if (this.state.isNewItemMode) {
      if (this.state.newItem)
        this.props.bkService.updateAreaRelationsInfo(this.state.newItem, 'Contoso').then(results => {
          console.log(results);
          this.setState({ itemsWithError: [], isLoadingConfirm: false, showDialog: false, hasErrorUpdating: false, isNewItemMode: false, newItem: undefined });
          this.loadData(true);
        }).catch(ex => {
          this.setState({ hasErrorUpdating: true, isLoadingConfirm: false, newItem: undefined, isNewItemMode: false })
          console.log(ex);
        })
    } else {
      const allpromises: promise<any>[] = [];
      const updatedItems = this.onGetUpdatedItems();
      updatedItems.forEach(item =>
        allpromises.push(this.props.bkService.updateAreaRelationsInfo(item, 'Contoso'))
      )

      promise.all(allpromises).then(results => {
        console.log(results);
        this.setState({ itemsWithError: [], isLoadingConfirm: false, showDialog: false, hasErrorUpdating: false });
        this.loadData(true);
      }).catch(ex => {
        this.setState({ hasErrorUpdating: true, isLoadingConfirm: false })
        console.log(ex);
      })
    }
  }

  private onClickCloseConfirmation = (): void => {
    this._toUpdateItems = [];
    if (this.state.hasErrorUpdating) {
      this.loadData(true);
    }
    this.setState({ showDialog: false, hasError: false, hasErrorUpdating: false, isNewItemMode: false })
  }

  private onClickNewItem = (): void => {
    this.setState({ isNewItemMode: true, showDialog: true, newItem: undefined });
  }

  public render(): React.ReactElement<IBusinessAreaRelationsprops> {
    const {
      title,
      isMobile
    } = this.props;
    const { filters, filteredRelations, isLoading, showDialog, itemsWithError, isLoadingConfirm, hasErrorUpdating, hasError, isNewItemMode, newItem } = this.state;

    let uniqueValues: string[] = [];
    if (showDialog && itemsWithError && itemsWithError.length > 0) {
      uniqueValues = [...new Set(itemsWithError.map(item => this._BusinessAreaDictionary[item.BusinessArea]))];
    }
    return (
      <div className={styles.BusinessAreaRelations}>
        {
          (title && title !== "") &&
          <div className={styles.wpTitle}>{this.props.title}</div>
        }
        {isLoading ?
          <div className={styles.resultZone}>
            <Spinner labelPosition="below" label={strings.IsLoadingItems} />
          </div>
          :
          <div className={styles.contentZone}>
            {hasError ?
              <div className={styles.filterZone}>
                {strings.ErrorUpdating}
              </div>
              :
              <>
                <div className={styles.filterZone}>
                  <Combobox
                    onOptionSelect={(_event, data): void => {
                      this.onChangeFilterComboBoxValue(RelationsFilterKeys.BusinessArea, data.optionValue)
                    }}
                    selectedOptions={filters && filters[RelationsFilterKeys.BusinessArea] ? filters[RelationsFilterKeys.BusinessArea] : []}
                    clearable
                    placeholder={strings.BusinessAreaPlaceholder}
                    className={styles.filterDropdown}
                  //value={filters && filters[RelationsFilterKeys.BusinessArea] && this._BusinessAreaDictionary[filters[RelationsFilterKeys.BusinessArea]]}
                  >
                    {this.onRenderOptions(RelationsFilterKeys.BusinessArea)}
                  </Combobox>
                  <Combobox
                    onOptionSelect={(_event, data): void => {
                      this.onChangeFilterComboBoxValue(RelationsFilterKeys.Division, data.optionValue)
                    }}
                    selectedOptions={filters && filters[RelationsFilterKeys.Division] ? filters[RelationsFilterKeys.Division] : []}
                    clearable
                    placeholder={strings.DivisionPlaceholder}
                    className={styles.filterDropdown}
                  //value={filters && filters[RelationsFilterKeys.Division] && this._ministryDictionary[filters[RelationsFilterKeys.Division]]}
                  >
                    {this.onRenderOptions(RelationsFilterKeys.Division)}
                  </Combobox>
                  <Switch
                    className={styles.filterSwitch}
                    checked={filters && !filters[RelationsFilterKeys.Actual]}
                    label={strings.HistoricoLabel}
                    onChange={(ev, data) => this.onChangeFilterBooleanValue(RelationsFilterKeys.Actual, !data.checked)}
                  />
                  <div className={styles.newElementButton}>
                    <Button icon={<AddRegular />}
                      onClick={() => this.onClickNewItem()}
                    >
                      {strings.NewElementLabel}
                    </Button>
                  </div>
                </div>
                <div className={styles.resultZone}>
                  <div className={styles.itemsZone}>
                    <DataGrid
                      items={filteredRelations}
                      columns={this.areaColumns}
                      className={styles.dataGridContainer}
                      noNativeElements={isMobile}
                    >
                      <DataGridHeader className={styles.dataGridHeadersContainer}>
                        <DataGridRow
                        >
                          {({ renderHeaderCell }) => (
                            <DataGridHeaderCell className={styles.dataGridHeaders}>{renderHeaderCell()}</DataGridHeaderCell>
                          )}
                        </DataGridRow>
                      </DataGridHeader>
                      <DataGridBody<IRelationBusinessArea> className={styles.dataGridBody}>
                        {({ item, rowId }) =>
                        (
                          <DataGridRow<IRelationBusinessArea>
                            key={rowId}
                            className={styles.dataRowItems}                >
                            {({ renderCell, columnId }) => (
                              <DataGridCell focusMode={"none"} className={isMobile && columnId === "openItem" ? styles.manageItem : ''}>
                                {renderCell(item)}
                              </DataGridCell>
                            )}
                          </DataGridRow>
                        )}
                      </DataGridBody>
                    </DataGrid>
                  </div>
                  <div className={styles.footerContainer}>
                    <Button
                      appearance="secondary"
                      onClick={() => this.onClickDiscard()}
                    >
                      {strings.DiscardLabel}
                    </Button>
                    <Button
                      appearance="primary"
                      onClick={() => this.onClickOpenDialog()}
                    >
                      {strings.SaveLabel}
                    </Button>
                  </div>
                </div>
              </>
            }
          </div>
        }
        {showDialog &&
          <Dialog open={showDialog}>
            <DialogSurface className={styles.relationsDialog}>
              <DialogBody>
                <DialogTitle>
                  {strings.DialogTitle}
                </DialogTitle>
                {isNewItemMode ?
                  <DialogContent>

                    {this.onRenderNewItemDialog()}
                  </DialogContent>
                  :
                  <>
                    {itemsWithError && itemsWithError.length > 0 ?
                      <DialogContent>
                        <div>{strings.ErrorItems}</div>
                        <ul>
                          {itemsWithError && uniqueValues && uniqueValues.map(item => {
                            return (
                              <li>{item}</li>
                            )
                          })}
                        </ul>
                      </DialogContent>
                      :
                      <DialogContent>
                        {isLoadingConfirm ?
                          <Spinner labelPosition="below" />
                          :
                          <div className={hasErrorUpdating ? styles.errorTxt : ''}>
                            {!hasErrorUpdating ? strings.ConfirmLabel : strings.ErrorUpdating}
                            {!hasErrorUpdating && this._toUpdateItems && this._toUpdateItems.length > 0 &&
                              <div className={styles.listChangesContainer}>
                                <TableHeader>
                                  <TableRow className={styles.listHeaderContainer}>
                                    <TableHeaderCell className={styles.listHeaderTitle}>
                                      {strings.BusinessAreaLabel}
                                    </TableHeaderCell>
                                    <TableHeaderCell className={styles.listHeaderTitle}>
                                      {strings.StartDateLabel}
                                    </TableHeaderCell >
                                    <TableHeaderCell className={styles.listHeaderTitle}>
                                      {strings.EndDateLabel}
                                    </TableHeaderCell>
                                    <TableHeaderCell className={styles.listHeaderTitle}>
                                      {strings.DivisionLabel}
                                    </TableHeaderCell>
                                  </TableRow>
                                </TableHeader>
                                {this._toUpdateItems.map(item => this.onRenderDifferences(item))}
                              </div>
                            }

                          </div>
                        }

                      </DialogContent>
                    }
                  </>
                }
                <DialogActions>
                  <DialogTrigger disableButtonEnhancement>
                    <Button
                      appearance="secondary"
                      onClick={() => this.onClickCloseConfirmation()}
                      disabled={isLoadingConfirm}
                    >
                      {strings.CancelLabel}
                    </Button>
                  </DialogTrigger>
                  {//!errorUpdating &&
                    <Button
                      appearance="primary"
                      onClick={() => this.onClickSaveConfirmation()}
                      disabled={
                        (itemsWithError && itemsWithError.length > 0)
                        || isLoadingConfirm
                        || hasErrorUpdating
                        || ((isNewItemMode) && (!newItem || !newItem[RelationAreaFields.BusinessArea] || !newItem[RelationAreaFields.StartDate] || !newItem[RelationAreaFields.EndDate] || !newItem[RelationAreaFields.Division] || newItem[RelationAreaFields.HasError]))
                      }
                    >
                      {strings.AcceptLabel}
                    </Button>
                  }
                </DialogActions >
              </DialogBody>
            </DialogSurface>
          </Dialog >

        }
      </div>
    );
  }

  private onRenderOptions = (key: RelationsFilterKeys): JSX.Element[] => {
    const { BusinessAreaFilterOptions, ministryFilterOptions, filteredRelations } = this.state;
    const { BusinessArea } = this.state.filters
    if (key === RelationsFilterKeys.BusinessArea) {
      let options: Term[] = [];
      options = BusinessAreaFilterOptions;
      return options.map((option: Term): JSX.Element => (
      <Option
        key={option.key}
        value={option.key}
      >
        {option.text}
      </Option>
    ))
    } else if (key === RelationsFilterKeys.Division) {
        let options: ILookupMinistryField[] = [];
      if (BusinessArea && BusinessArea.length > 0) {

        options = ministryFilterOptions.filter(min => filteredRelations.some(element => min.Id.toString() == element.Division.toString()));
      } else {
        options = ministryFilterOptions
      }

      return options.map((option: ILookupMinistryField): JSX.Element => (
        <Option
          key={option.Id}
          value={option.Id}
        >
          {option.Title}
        </Option>
    ))
    }
    else{
      const elementos: JSX.Element[] = [];
      return elementos;
    }
  }

  private onRenderDifferences = (item: IRelationBusinessArea): JSX.Element => {
    /*return (
      <li className={styles.listRowChange}>
        <div>
          <Label className={styles.areaChange}>{this._BusinessAreaDictionary[item[RelationAreaFields.BusinessArea]]}</Label> -
          <Label className={styles.valueChange}>{item?.StartDate ? format(new Date(item[RelationAreaFields.StartDate]), 'dd-MM-yyyy') : "(vacio)"}</Label> -
          <Label className={styles.valueChange}>{item?.EndDate ? format(new Date(item[RelationAreaFields.EndDate]), 'dd-MM-yyyy') : "(vacio)"}</Label> -
          <Label className={styles.valueChange}>{this._ministryDictionary[item[RelationAreaFields.Division]]}</Label>
        </div>
      </li>)*/

    return (
      <TableRow key={item[RelationAreaFields.Id]} className={styles.listRowChange}>
        <TableCell className={styles.areaChange}>
          {this._BusinessAreaDictionary[item[RelationAreaFields.BusinessArea]]}
        </TableCell>
        <TableCell>
          <TableCellLayout>
            {item?.[RelationAreaFields.StartDate] ? format(new Date(item[RelationAreaFields.StartDate]), 'dd-MM-yyyy') : "(" + strings.ECNTyValue + ")"}
          </TableCellLayout>
        </TableCell>
        <TableCell>
          {item?.[RelationAreaFields.EndDate] ? format(new Date(item[RelationAreaFields.EndDate]), 'dd-MM-yyyy') : "(" + strings.ECNTyValue + ")"}
        </TableCell>
        <TableCell>
          {this._ministryDictionary[item[RelationAreaFields.Division]]}
        </TableCell>
      </TableRow>
    )
  }

  private onRenderNewItemDialog = (): JSX.Element => {
    const { newItem, isLoadingConfirm, hasErrorUpdating } = this.state;
    console.log(newItem && newItem[RelationAreaFields.StartDate] && format(new Date(newItem[RelationAreaFields.StartDate]), 'yyyy-MM-dd'))
    console.log(newItem?.[RelationAreaFields.EndDate] && format(new Date(newItem[RelationAreaFields.EndDate]), 'yyyy-MM-dd'))

    const diffItems = this.onGetUpdatedItems();
    if (diffItems && diffItems.length > 0) {
      return (
        <div className={styles.errorTxt}>
          {strings.ReviewPendingChanges}
        </div>)
    } else {
      return (
        <div className={styles.newItemForm}>
          {isLoadingConfirm ?
            <Spinner labelPosition="below" />
            :
            hasErrorUpdating ?
              <div className={styles.errorTxt}>
                {strings.ErrorUpdating}
              </div>
              :
              <>
                <div className={styles.newItemFieldFull}>
                  <Label className={styles.newItemLabel}>{strings.BusinessAreaLabel}</Label>
                  <Combobox
                    onOptionSelect={(_event, data): void => {
                      data.optionValue && this.onUpdateValue(RelationAreaFields.BusinessArea, undefined, data.optionValue)
                    }}
                    selectedOptions={newItem && newItem[RelationsFilterKeys.BusinessArea] ? [newItem[RelationsFilterKeys.BusinessArea]] : []}
                    placeholder={strings.BusinessAreaPlaceholder}
                    className={styles.filterDropdown}
                    value={newItem && newItem[RelationAreaFields.BusinessArea] && this._BusinessAreaDictionary[newItem[RelationAreaFields.BusinessArea]]}
                    title={newItem && newItem[RelationAreaFields.BusinessArea] && this._BusinessAreaDictionary[newItem[RelationAreaFields.BusinessArea]]}
                  >
                    {this.onRenderOptions(RelationsFilterKeys.BusinessArea)}
                  </Combobox>
                </div>
                <div className={styles.newItemHalfField}>
                  <div className={styles.newItemDate} key={'StartInit'}>
                    <Label className={styles.newItemLabel}>{strings.StartDateLabel}</Label>
                    <Input
                      key='StartNewDate'
                      value={newItem && newItem[RelationAreaFields.StartDate] && format(new Date(newItem[RelationAreaFields.StartDate]), 'yyyy-MM-dd')}
                      type='date'
                      onChange={async (data): promise<void> => {
                        this.onUpdateValue(RelationAreaFields.StartDate, undefined, data.target.value ? new Date(data.target.value).toISOString() : "")
                      }}
                      onKeyDown={(e): void => e.preventDefault()}
                      className={styles.dateSelector}
                      title={newItem?.[RelationAreaFields.StartDate] && format(new Date(newItem[RelationAreaFields.StartDate]), 'dd-MM-yyyy')}
                    />
                  </div>
                  <div className={styles.newItemDate} key={'EndInit'}>
                    <Label className={styles.newItemLabel}>{strings.EndDateLabel}</Label>
                    <Input
                      key='EndNewDate'
                      type='date'
                      onChange={async (data): promise<void> => {
                        this.onUpdateValue(RelationAreaFields.EndDate, undefined, data.target.value ? new Date(data.target.value).toISOString() : "")
                      }}
                      onKeyDown={(e): void => e.preventDefault()}
                      className={styles.dateSelector}
                      title={newItem?.[RelationAreaFields.EndDate] && format(new Date(newItem[RelationAreaFields.EndDate]), 'dd-MM-yyyy')}
                      value={newItem?.[RelationAreaFields.EndDate] && format(new Date(newItem[RelationAreaFields.EndDate]), 'dd-MM-yyyy')}
                      min={newItem?.StartDate && format(new Date(newItem.StartDate), 'yyyy-MM-dd')}
                    />
                  </div>
                </div>
                <div className={styles.newItemFieldFull}>
                  <Label className={styles.newItemLabel}>{strings.DivisionLabel}</Label>
                  <Combobox
                    onOptionSelect={(_event, data): void => {
                      data.optionValue && this.onUpdateValue(RelationAreaFields.Division, undefined, data.optionValue)
                    }}
                    selectedOptions={newItem && newItem[RelationsFilterKeys.Division] ? [newItem[RelationsFilterKeys.Division]] : []}
                    placeholder={strings.DivisionPlaceholder}
                    className={styles.filterDropdown}
                    value={newItem && newItem[RelationAreaFields.Division] && this._ministryDictionary[newItem[RelationAreaFields.Division]]}
                    title={newItem && newItem[RelationAreaFields.Division] && this._ministryDictionary[newItem[RelationAreaFields.Division]]}
                  >
                    {this.onRenderOptions(RelationsFilterKeys.Division)}
                  </Combobox>
                </div>
                <div className={styles.newItemErrorRow}>
                  {newItem && newItem[RelationAreaFields.HasError] &&
                    <Label>{strings.ErrorNewItem}</Label>
                  }
                </div>
              </>
          }
        </div>
      )
    }
  }
}
