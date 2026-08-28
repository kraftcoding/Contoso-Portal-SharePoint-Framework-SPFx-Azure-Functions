import * as React from 'react';
import styles from './ResultZone.module.scss';
import {
  IResultZoneprops,
  IResultZoneState
} from './IResultZone';
import { IOrganEXTERNALResult } from '../../models/CertificationAppModels';
import { Button, createTableColumn, DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell, DataGridprops, DataGridRow, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner, TableCellLayout, TableColumnDefinition, Tooltip } from '@fluentui/react-components';
import { ChevronLeftFilled, ChevronRightFilled, DocumentRibbonRegular, EditSettingsRegular, FolderRegular, SettingsRegular } from '@fluentui/react-icons';
import _ from 'lodash';
import strings from 'CertificationAppWebPartStrings';
class ResultZone extends React.Component<IResultZoneprops, IResultZoneState> {

  private loadingDowndloadCertificateBodyId: string | undefined = undefined;

  private columnsOrgan: TableColumnDefinition<IOrganEXTERNALResult>[] =
    this.props.isMobile ?
    [
      createTableColumn<IOrganEXTERNALResult>({
        columnId: "title",
        renderHeaderCell: () => {
          return strings.NameHeader;
        },
        renderCell: (item) => {
          return (
            <TableCellLayout className={styles.dataItemTitle}>
              {this.props?.organListDictionary[item.Department]}
            </TableCellLayout>
          );
        },
      }),
      createTableColumn<IOrganEXTERNALResult>({
        columnId: "openItem",
        renderHeaderCell: () => {
          return "";
        },
        renderCell: (item) => {
          return (<div className={styles.manageItem}>
            <Menu>
              <MenuTrigger>
                {this?.state?.isDownloadingCertificate ?
                  this.loadingDowndloadCertificateBodyId === item.BodyId ?
                    <Spinner />
                    :
                    <Button disabled appearance={"secondary"} icon={<SettingsRegular />} />
                  :
                  <Button appearance={"secondary"} icon={<SettingsRegular />} />
                }
              </MenuTrigger>
              <MenuPopover>
                <MenuList>
                  <MenuItem icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.CodigoDepartment, "Open")}>{strings.ManageLabel}</MenuItem>
                  <MenuItem icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.CodigoDepartment)}>{strings.CertificateLabel}</MenuItem>
                  <MenuItem icon={<FolderRegular />} onClick={() => window.open(item.FileRef,'_blank')}>{strings.DocumentsTooltip}</MenuItem>
                </MenuList>
              </MenuPopover>
            </Menu>
          </div>)
        },
      })
    ]
    :
    [
      createTableColumn<IOrganEXTERNALResult>({
        columnId: "title",
        compare: (a, b) => {
          return this.props?.organListDictionary[a.Department]?.localeCompare(this.props?.organListDictionary[b.Department]);
        },
        renderHeaderCell: () => {
          return strings.NameHeader;
        },
        renderCell: (item) => {
          return (
            <TableCellLayout>
              {this.props?.organListDictionary[item.Department]}
            </TableCellLayout>
          );
        },
      }),
      createTableColumn<IOrganEXTERNALResult>({
        columnId: "type",
        compare: (a, b) => {
          return this.props?.organTypeListDictionary[a.TipoDepartmentEXTERNAL]?.localeCompare(this.props?.organTypeListDictionary[b.TipoDepartmentEXTERNAL]);
        },
        renderHeaderCell: () => {
          return strings.TypeHeader;
        },
        renderCell: (item) => {
          return (
            <TableCellLayout
              truncate
            >
              {this.props?.organTypeListDictionary[item.TipoDepartmentEXTERNAL]}
            </TableCellLayout>
          );
        },
      }),
      createTableColumn<IOrganEXTERNALResult>({
        columnId: "openItem",
        renderHeaderCell: () => {
          return "";
        },
        renderCell: (item) => {
          return (
            <div className={styles.buttonActions}>
              <Tooltip
                withArrow
                positioning={'below'}
                content={strings.ManageLabel}
                relationship="label"
              >
                <Button appearance={"secondary"} icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.CodigoDepartment, "Open")} />
              </Tooltip>
              <Tooltip
                withArrow
                positioning={'below'}
                content={strings.CertificateLabel}
                relationship="label"
              >
                {this?.state?.isDownloadingCertificate ?
                  this.loadingDowndloadCertificateBodyId === item.CodigoDepartment ?
                    <Spinner />
                    :
                    <Button disabled={this.state.isDownloadingCertificate} appearance={"secondary"} icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.CodigoDepartment)} />
                  :
                  <Button appearance={"secondary"} icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.CodigoDepartment)} />
                }
              </Tooltip>
              <Tooltip
                withArrow
                positioning={'below'}
                content={strings.DocumentsTooltip}
                relationship="label"
              >
                <Button appearance={"secondary"} icon={<FolderRegular />} onClick={() => window.open(item.FileRef,'_blank')} />
              </Tooltip>
            </div>
          )
        },
      })
    ];

  constructor(props: IResultZoneprops) {
    super(props);
    this.state = {
      sorting: { sortColumn: '', sortDirection: 'ascending' },
      pagedBackOrgans: [],
      isDownloadingCertificate: false
    };
  }

  componentDidMount(): void {
    this.setState({
      pagedBackOrgans: this.props.organBackList.slice(0, this.props.numberPageItems),
      //sorting: this.props.searchType === SearchType.Organ ? {sortColumn: 'title', sortDirection: 'ascending' } : this.props.searchType === SearchType.User ? { sortColumn: '', sortDirection: 'ascending' } : { sortColumn: '', sortDirection: 'ascending' }, 
      sorting: { sortColumn: 'title', sortDirection: 'ascending' },
    })
  }

  componentDidUpdate(testvprops: Readonly<IResultZoneprops>): void {
    if (!_.isEqual(this.props.organBackList, testvprops.organBackList)) {
      this.setState({ pagedBackOrgans: this.props.organBackList.slice(0, this.props.numberPageItems) })
    }
    //sorting: { sortColumn: 'title', sortDirection: 'ascending' },

    if (this.props.currentPage !== testvprops.currentPage && _.isEqual(this.props.organBackList, testvprops.organBackList)) {
      const orderedItems = this.getSortedItems(this.state.sorting);
      this.setState({ pagedBackOrgans: orderedItems.slice((this.props.currentPage - 1) * this.props.numberPageItems, this.props.currentPage * this.props.numberPageItems) as IOrganEXTERNALResult[] })
    }
  }

  private onClickDownloadCertificate(bodyId: string, isEXTERNAL?: boolean) {
    console.log(bodyId);
    console.log(isEXTERNAL);
    this.loadingDowndloadCertificateBodyId = bodyId;
    this.setState({ isDownloadingCertificate: true });
    this.props.onDownloadCertificate(bodyId, true).then(result => {
      if (result) {
        const url = window.URL.createObjectURL(new Blob([result], { type: "application/pdf" }));
        const enlace = document.createElement('a');
        enlace.href = url;
        enlace.setAttribute('download', `${bodyId} - certificate.pdf`);
        document.body.appendChild(enlace);
        enlace.click();
        document.body.removeChild(enlace);
        window.URL.revokeObjectURL(url);
      }
      this.loadingDowndloadCertificateBodyId = undefined;
      this.setState({ isDownloadingCertificate: false });
    }
    ).catch(ex => {
      console.log(ex)
      this.loadingDowndloadCertificateBodyId = undefined;
      this.setState({ isDownloadingCertificate: false });
    }
    )
  }

  private onSort: DataGridprops["onSortChange"] = (e, nextSortState) => {
    const orderedItems = this.getSortedItems(nextSortState);
    this.setState({ sorting: nextSortState, pagedBackOrgans: orderedItems.slice(0, this.props.numberPageItems) as IOrganEXTERNALResult[] });
  
    this.props.onUpdatePage(1);
  }

  private getSortedItems(sort: DataGridprops["sortState"]): (IOrganEXTERNALResult[]) {
      let orderedItems: IOrganEXTERNALResult[] = [...this.props.organBackList];
      if (sort !== undefined) {
        if (sort.sortColumn === "title") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => this.props?.organListDictionary[a.Department]?.localeCompare(this.props?.organListDictionary[b.Department]) : (a, b) => this.props?.organListDictionary[b.Department]?.localeCompare(this.props?.organListDictionary[a.Department]));
        } else if (sort.sortColumn === "naming") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => a.BodyId?.localeCompare(b.BodyId) : (a, b) => b.BodyId?.localeCompare(a.BodyId));
        } else if (sort.sortColumn === "type") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => this.props?.organTypeListDictionary[a.TipoDepartmentEXTERNAL]?.localeCompare(this.props?.organTypeListDictionary[b.TipoDepartmentEXTERNAL]) : (a, b) => this.props?.organTypeListDictionary[b.TipoDepartmentEXTERNAL]?.localeCompare(this.props?.organTypeListDictionary[a.TipoDepartmentEXTERNAL]));
        }
      }
      return orderedItems;
  }

  private setPage(page: number): void {
    this.props.onUpdatePage(page);
  }

  public render(): JSX.Element {
    const { organBackList,  isMobile, numberPageItems, currentPage } = this.props;
    const { pagedBackOrgans } = this.state;
    return (
      <div
        className={styles.resultZone}
      >
        {organBackList && organBackList.length > 0 ?
          <DataGrid
            items={ pagedBackOrgans}
            columns={this.columnsOrgan}
            className={styles.dataGridContainer}
            noNativeElements={isMobile}
            sortable
            onSortChange={this.onSort}
            sortState={this.state.sorting}
          >
            <DataGridHeader className={styles.dataGridHeadersContainer}>
              <DataGridRow
              >
                {({ renderHeaderCell }) => (
                  <DataGridHeaderCell className={styles.dataGridHeaders}>{renderHeaderCell()}</DataGridHeaderCell>
                )}
              </DataGridRow>
            </DataGridHeader>
            <DataGridBody<IOrganEXTERNALResult> className={styles.dataGridBody}>
              {({ item, rowId }) =>
              (
                <DataGridRow<IOrganEXTERNALResult>
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
          :
          <div className={styles.noResults}>{strings.NoResults}</div>
        }
        {numberPageItems < organBackList.length &&
          <div className={styles.pagingContainer}>
            <Button
              //className={styles.searchButton}
              onClick={() => this.setPage(currentPage - 1)}
              appearance='secondary'
              disabled={currentPage <= 1}
              icon={<ChevronLeftFilled />}
            />
            <label className={styles.pageNumber}>{strings.PageNumberTxt} {currentPage} {strings.FromTotalPageTxt} {Math.ceil(organBackList.length / numberPageItems)}</label>
            <Button
              //className={styles.searchButton}
              onClick={() => this.setPage(currentPage + 1)}
              appearance='secondary'
              disabled={(currentPage * numberPageItems) >= organBackList.length}
              icon={<ChevronRightFilled />}
            />

          </div>
        }
      </div>
    );
  }
  /*
  private onManageItem(id: string): void {
    console.log(id);
  }
    */
}

export default ResultZone;