import * as React from 'react';
import styles from './ResultZone.module.scss';
import {
  IResultZoneprops,
  IResultZoneState
} from './IResultZone';
import { IOrganAdministration, IOrganResult, IUserResult, LicenseNames, LicensesTypes, SearchType } from '../../models/AdministrationAppModels';
import { Button, createTableColumn, DataGrid, DataGridBody, DataGridCell, DataGridHeader, DataGridHeaderCell, DataGridprops, DataGridRow, Link, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner, TableCellLayout, TableColumnDefinition, Tooltip } from '@fluentui/react-components';
import { ChevronLeftFilled, ChevronRightFilled, DocumentRibbonRegular, EditSettingsRegular, PersonEditRegular, SettingsRegular } from '@fluentui/react-icons';
import _ from 'lodash';
import strings from 'AdministrationAppWebPartStrings';
class ResultZone extends React.Component<IResultZoneprops, IResultZoneState> {

  private loadingDowndloadCertificateBodyId: string | undefined = undefined;

  private columnsOrgan: TableColumnDefinition<IOrganResult>[] =
    this.props.isMobile ?
      this.props.showEXTERNAL ?
        [
          createTableColumn<IOrganResult>({
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
          createTableColumn<IOrganResult>({
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
                      <MenuItem icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.BodyId, this.props.searchType, "Open")}>{strings.ManageLabel}</MenuItem>
                      <MenuItem icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.BodyId, item.IsEXTERNAL)}>{strings.CertificateLabel}</MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </div>)
            },
          })
        ]
        :
        [
          createTableColumn<IOrganResult>({
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
          createTableColumn<IOrganResult>({
            columnId: "openItem",
            renderHeaderCell: () => {
              return "";
            },
            renderCell: (item) => {
              return (<div className={styles.manageItem}>
                <Menu>
                  <MenuTrigger>
                    <Button appearance={"secondary"} icon={<SettingsRegular />} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.BodyId, this.props.searchType, "Open")}>{strings.ManageLabel}</MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </div>)
            },
          })
        ]
      :
      this.props.showEXTERNAL ?
        [
          createTableColumn<IOrganResult>({
            columnId: "title",
            compare: (a, b) => {
              return this.props?.organListDictionary[a.Department].localeCompare(this.props?.organListDictionary[b.Department]);
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
          createTableColumn<IOrganResult>({
            columnId: "type",
            compare: (a, b) => {
              return this.props?.organTypeListDictionary[a.TipoDepartment].localeCompare(this.props?.organTypeListDictionary[b.TipoDepartment]);
            },
            renderHeaderCell: () => {
              return strings.TypeHeader;
            },
            renderCell: (item) => {
              return (
                <TableCellLayout
                  truncate
                >
                  {this.props?.organTypeListDictionary[item.TipoDepartment]}
                </TableCellLayout>
              );
            },
          }),
          createTableColumn<IOrganResult>({
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
                    <Button appearance={"secondary"} icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.BodyId, this.props.searchType, "Open")} />
                  </Tooltip>
                  <Tooltip
                    withArrow
                    positioning={'below'}
                    content={strings.CertificateLabel}
                    relationship="label"
                  >
                    {this?.state?.isDownloadingCertificate ?
                      this.loadingDowndloadCertificateBodyId === item.BodyId ?
                        <Spinner />
                        :
                        <Button disabled={this.state.isDownloadingCertificate} appearance={"secondary"} icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.BodyId, item.IsEXTERNAL)} />
                      :
                      <Button appearance={"secondary"} icon={<DocumentRibbonRegular />} onClick={() => this.onClickDownloadCertificate(item.BodyId, item.IsEXTERNAL)} />
                    }
                  </Tooltip>
                </div>
              )
            },
          })
        ]
        :
        [
          createTableColumn<IOrganResult>({
            columnId: "title",
            compare: (a, b) => {
              return this.props?.organListDictionary[a.Department].localeCompare(this.props?.organListDictionary[b.Department]);
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
          createTableColumn<IOrganResult>({
            columnId: "naming",
            compare: (a, b) => {
              return a.BodyId.localeCompare(b.BodyId);
            },
            renderHeaderCell: () => {
              return strings.NomenclaturaHeader;
            },
            renderCell: (item) => {
              return (
                <TableCellLayout
                >
                  <Link href={`/sites/${item.BodyId}`}>{item.BodyId}</Link>
                </TableCellLayout>
              );
            },
          }),
          createTableColumn<IOrganResult>({
            columnId: "type",
            compare: (a, b) => {
              return this.props?.organTypeListDictionary[a.TipoDepartment].localeCompare(this.props?.organTypeListDictionary[b.TipoDepartment]);
            },
            renderHeaderCell: () => {
              return strings.TypeHeader;
            },
            renderCell: (item) => {
              return (
                <TableCellLayout
                  truncate
                >
                  {this.props?.organTypeListDictionary[item.TipoDepartment]}
                </TableCellLayout>
              );
            },
          }),
          createTableColumn<IOrganResult>({
            columnId: "openItem",
            renderHeaderCell: () => {
              return "";
            },
            renderCell: (item) => {
              return (<Tooltip
                withArrow
                positioning={'below'}
                content={strings.ManageLabel}
                relationship="label"
              >
                <Button appearance={"secondary"} icon={<EditSettingsRegular />} onClick={() => this.props.onManageItem(item.BodyId, this.props.searchType, "Open")} />
              </Tooltip>)
            },
          })
        ];

  private columnsUser: TableColumnDefinition<IUserResult>[] =
    this.props.isMobile ?
      [
        createTableColumn<IUserResult>({
          columnId: "title",
          renderHeaderCell: () => {
            return strings.NameHeader;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout className={styles.dataItemTitle}>
                {item.DisplayName}
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IUserResult>({
          columnId: "openItem",
          renderHeaderCell: () => {
            return "";
          },
          renderCell: (item) => {
            return (
              <div className={styles.manageItem}>
                <Menu>
                  <MenuTrigger>
                    <Button appearance={"secondary"} icon={<SettingsRegular />} />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem icon={<PersonEditRegular />} onClick={() => this.props.onManageItem(item.UserPrincipalName, this.props.searchType, "Open")}>{strings.ManageLabel}</MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              </div>)
          },
        })
      ] :
      [
        createTableColumn<IUserResult>({
          columnId: "title",
          compare: (a, b) => {
            return a.DisplayName.localeCompare(b.DisplayName);
          },
          renderHeaderCell: () => {
            return strings.NameHeader;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout>
                {item.DisplayName}
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IUserResult>({
          columnId: "mail",
          compare: (a, b) => {
            return a.PrincipalMail.localeCompare(b.PrincipalMail);
          },
          renderHeaderCell: () => {
            return strings.MailHeader;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout
              >
                {item.PrincipalMail}
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IUserResult>({
          columnId: "organ",
          renderHeaderCell: () => {
            return strings.OrgansHeader;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout
              >
                <div>
                  {item.Organ && item.Organ.map((org) =>
                    <div className={styles.userOrgansList}>
                      {org.DepartmentName} - <label className={styles.organListRole}>{org.Role}</label>
                    </div>
                  )}
                </div>
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IUserResult>({
          columnId: "license",
          renderHeaderCell: () => {
            return strings.LicenseFilter;
          },
          renderCell: (item) => {
            return (
              <TableCellLayout
              >
                <div>
                  {item.AssignedLicenses && item.AssignedLicenses.filter(lic => lic === LicensesTypes.E3 || lic === LicensesTypes.E5 || lic === LicensesTypes.E5Developer).map((lic:string) =>
                    <div className={styles.userOrgansList}>
                      <label className={styles.organListRole}>{lic === LicensesTypes.E3 ? LicenseNames[LicensesTypes.E3] : LicenseNames[LicensesTypes.E5]}</label>
                    </div>
                  )}
                </div>
              </TableCellLayout>
            );
          },
        }),
        createTableColumn<IUserResult>({
          columnId: "openItem",
          renderHeaderCell: () => {
            return "";
          },
          renderCell: (item) => {
            return (<Tooltip
              withArrow
              positioning={'below'}
              content={strings.ManageLabel}
              relationship="label"
            >
              <Button appearance={"secondary"} icon={<PersonEditRegular />} onClick={() => this.props.onManageItem(item.UserPrincipalName, this.props.searchType, "Open")} />
            </Tooltip>)
          },
        })
      ]
    ;

  constructor(props: IResultZoneprops) {
    super(props);
    this.state = {
      sorting: { sortColumn: '', sortDirection: 'ascending' },
      pagedBackOrgans: [],
      pagedBackUsers: [],
      isDownloadingCertificate: false
    };
  }

  componentDidMount(): void {
    this.setState({
      pagedBackOrgans: this.props.organBackList.slice(0, this.props.numberPageItems),
      pagedBackUsers: this.props.userBackList.slice(0, this.props.numberPageItems),
      //sorting: this.props.searchType === SearchType.Organ ? {sortColumn: 'title', sortDirection: 'ascending' } : this.props.searchType === SearchType.User ? { sortColumn: '', sortDirection: 'ascending' } : { sortColumn: '', sortDirection: 'ascending' }, 
      sorting: { sortColumn: 'title', sortDirection: 'ascending' },
    })
  }

  componentDidUpdate(testvprops: Readonly<IResultZoneprops>): void {
    if (!_.isEqual(this.props.organBackList, testvprops.organBackList)) {
      this.setState({ pagedBackOrgans: this.props.organBackList.slice(0, this.props.numberPageItems) })
    }
    if (!_.isEqual(this.props.userBackList, testvprops.userBackList)) {
      this.setState({ pagedBackUsers: this.props.userBackList.slice(0, this.props.numberPageItems) })
    }
    if (this.props.searchType !== testvprops.searchType) {
      this.setState({
        pagedBackOrgans: this.props.organBackList.slice(0, this.props.numberPageItems),
        pagedBackUsers: this.props.userBackList.slice(0, this.props.numberPageItems),
        sorting: { sortColumn: 'title', sortDirection: 'ascending' },
      })
    }

    if (this.props.currentPage !== testvprops.currentPage && this.props.searchType === testvprops.searchType && _.isEqual(this.props.userBackList, testvprops.userBackList) && _.isEqual(this.props.organBackList, testvprops.organBackList)) {
      const orderedItems = this.getSortedItems(this.state.sorting, this.props.searchType);
      if (this.props.searchType === SearchType.Organ) {
        this.setState({ pagedBackOrgans: orderedItems.slice((this.props.currentPage - 1) * this.props.numberPageItems, this.props.currentPage * this.props.numberPageItems) as IOrganResult[] })
      } else {
        this.setState({ pagedBackUsers: orderedItems.slice((this.props.currentPage - 1) * this.props.numberPageItems, this.props.currentPage * this.props.numberPageItems) as IUserResult[] })
      }
    }
  }

  private onClickDownloadCertificate(bodyId: string, isEXTERNAL?: boolean) {
    console.log(bodyId);
    console.log(isEXTERNAL);
    this.setState({ isDownloadingCertificate: true });
    this.loadingDowndloadCertificateBodyId = bodyId;
    this.props.onDownloadCertificate(bodyId, isEXTERNAL).then(result => {
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
    const orderedItems = this.getSortedItems(nextSortState, this.props.searchType);
    if (this.props.searchType === SearchType.Organ) {
      this.setState({ sorting: nextSortState, pagedBackOrgans: orderedItems.slice(0, this.props.numberPageItems) as IOrganResult[] });
    } else {
      this.setState({ sorting: nextSortState, pagedBackUsers: orderedItems.slice(0, this.props.numberPageItems) as IUserResult[] });
    }
    this.props.onUpdatePage(1);
  }

  private getSortedItems(sort: DataGridprops["sortState"], type: SearchType): (IOrganResult[] | IUserResult[]) {
    if (type === SearchType.Organ) {
      let orderedItems: IOrganResult[] = [...this.props.organBackList];
      if (sort !== undefined) {
        if (sort.sortColumn === "title") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => this.props?.organListDictionary[a.Department].localeCompare(this.props?.organListDictionary[b.Department]) : (a, b) => this.props?.organListDictionary[b.Department].localeCompare(this.props?.organListDictionary[a.Department]));
        } else if (sort.sortColumn === "naming") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => a.BodyId.localeCompare(b.BodyId) : (a, b) => b.BodyId.localeCompare(a.BodyId));
        } else if (sort.sortColumn === "type") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => this.props?.organTypeListDictionary[a.TipoDepartment].localeCompare(this.props?.organTypeListDictionary[b.TipoDepartment]) : (a, b) => this.props?.organTypeListDictionary[b.TipoDepartment].localeCompare(this.props?.organTypeListDictionary[a.TipoDepartment]));
        }
      }
      return orderedItems;

    } else {
      let orderedItems: IUserResult[] = [...this.props.userBackList];
      if (sort !== undefined) {
        if (sort.sortColumn === "title") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => a.DisplayName.localeCompare(b.DisplayName) : (a, b) => b.DisplayName.localeCompare(a.DisplayName));
        } else if (sort.sortColumn === "mail") {
          orderedItems.sort(sort.sortDirection === 'ascending' ? (a, b) => a.PrincipalMail.localeCompare(b.PrincipalMail) : (a, b) => b.PrincipalMail.localeCompare(a.PrincipalMail));
        }
      }
      return orderedItems;
    }
  }

  private setPage(page: number): void {
    this.props.onUpdatePage(page);
  }

  public render(): JSX.Element {
    const { organBackList, userBackList, searchType, isMobile, numberPageItems, currentPage } = this.props;
    const { pagedBackOrgans, pagedBackUsers } = this.state;
    return (
      <div
        className={styles.resultZone}
      >
        {(searchType === SearchType.Organ && organBackList && organBackList.length > 0) || (searchType === SearchType.User && userBackList && userBackList.length > 0) ?
          <DataGrid
            items={searchType === SearchType.Organ ? pagedBackOrgans : pagedBackUsers}
            columns={searchType === SearchType.Organ ? this.columnsOrgan : this.columnsUser}
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
            <DataGridBody<IOrganAdministration> className={styles.dataGridBody}>
              {({ item, rowId }) =>
              (
                <DataGridRow<IOrganAdministration>
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
        {searchType === SearchType.Organ && numberPageItems < organBackList.length &&
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
        {searchType === SearchType.User && numberPageItems < userBackList.length &&
          <div className={styles.pagingContainer}>
            <Button
              //className={styles.searchButton}
              onClick={() => this.setPage(currentPage - 1)}
              appearance='secondary'
              disabled={currentPage <= 1}
              icon={<ChevronLeftFilled />}
            />
            <label className={styles.pageNumber}>{strings.PageNumberTxt} {currentPage} {strings.FromTotalPageTxt} {Math.ceil(userBackList.length / numberPageItems)}</label>
            <Button
              //className={styles.searchButton}
              onClick={() => this.setPage(currentPage + 1)}
              appearance='secondary'
              disabled={(currentPage * numberPageItems) >= userBackList.length}
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