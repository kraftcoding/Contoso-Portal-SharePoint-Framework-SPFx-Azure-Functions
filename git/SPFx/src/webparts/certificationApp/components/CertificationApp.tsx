import * as React from 'react';
import styles from './CertificationApp.module.scss';
import type { ICertificationAppprops, ICertificationAppState } from './ICertificationApp';
import { Certification, ILookupEXTERNALField, IOrganEXTERNALResult, ISearchFilters, OrganTaxonomyIds, EXTERNALList} from '../models/CertificationAppModels';
import { Term } from '../../../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Button, Dialog, DialogBody, DialogSurface, DialogTitle, Divider, Spinner } from '@fluentui/react-components';
import { AddRegular, DocumentAddRegular, PersonAddRegular } from '@fluentui/react-icons';
import strings from 'CertificationAppWebPartStrings';
import SearchBar from './SearchBar/SearchBar';
import FilterZone from './FilterZone/FilterZone';
import ResultZone from './ResultZone/ResultZone';
import EXTERNALOrganDetail from './EXTERNALOrganDetail/EXTERNALOrganDetail';
import _ from 'lodash';
import EXTERNALDocumentForm from '../../../components/EXTERNALDocumentForm/EXTERNALDocumentForm';
import { IRelationBusinessArea } from '../../BusinessAreaRelations/components/BusinessAreaRelationsModels';
import { serializeError } from '../../../utils/errorUtils';

export default class CertificationApp extends React.Component<ICertificationAppprops, ICertificationAppState> {
  private _organListDictionary: Record<string, string>;
  private _organTypeListDictionary: Record<string, string>;
  private _ministryDictionary: Record<string, string>;
  private _allOrganCodes: string[];
  private _BusinessAreaDivisionRelations: IRelationBusinessArea[];
  private _documentTypesDictionary: Record<string, string>;

  constructor(props: ICertificationAppprops) {
    super(props);
    this.state = {
      searchText: "",
      organTypeFilterOptions: [],
      ministryFilterOptions: [],
      organFilterOptions: [],
      roleFilterOptions: [],
      BusinessAreaOptions: [],
      StatusFilterOptions: [],
      searchFilters: Certification.eCNTyFilters,
      enableRemoveFilters: false,
      dialogIsOpen: false,
      managedItem: {
        itemId: "",
        organDialogContent: undefined,
      },
      currentPage: 1,
      isLoading: true,
      organBackList: [],
      organBackResultList: [],
      showLoadDataError: false,
      loadDataError: null,
      isLoadingUpdates: true,
      errorUpdating: false,
      documentDialog: false
    };
  }

  async componentDidMount(): promise<void> {
    await this.getTaxonomies();
    void this.loadData();

  }
  private loadData(): void {

    this.props.bkService.getEXTERNALOrgansInfo("Contoso").then(results => {
      const organs: IOrganEXTERNALResult[] = results;
      const orderedOrgans = organs.sort((a, b) => this._organListDictionary[a.Department]?.localeCompare(this._organListDictionary[b.Department]));
      const organCodes:string[] = [];
      orderedOrgans.forEach(org => {
        org.NombreDepartment = this._organListDictionary[org.Department];
        organCodes.push(org.CodigoDepartment)
      });
      console.log(orderedOrgans);
      this._allOrganCodes=[...organCodes];
      this.setState({ organBackList: orderedOrgans, organBackResultList: orderedOrgans, isLoading: false });
    }).catch(ex => {
      this.setState({ showLoadDataError: true, isLoading: false, loadDataError: serializeError(ex) });
      console.error(ex);
    });
  }

  private async getTaxonomies(): promise<void> {
    const { spService, bkService } = this.props;

    const [organsTerms, organsTypeLookup, ministryLookup, roleTerms, areaSecTerms, StatusTerms, BusinessAreaRelation, documentTypes]: [ITermInfo[], ILookupEXTERNALField[], ILookupEXTERNALField[], ITermInfo[], ITermInfo[], ILookupEXTERNALField[], IRelationBusinessArea[], ILookupEXTERNALField[]] = await promise.all([
      spService.getTaxonomy(OrganTaxonomyIds.Organ),
      spService.getItems<ILookupEXTERNALField>(`${this.props.context.pageContext.site.serverRelativeUrl}${EXTERNALList.ListTipoDepartmentUrl}`, EXTERNALList.getAll()),
      spService.getItems<ILookupEXTERNALField>(`${this.props.context.pageContext.site.serverRelativeUrl}${EXTERNALList.ListDivisionUrl}`, EXTERNALList.getAll()),
      spService.getTaxonomy(OrganTaxonomyIds.Roles),
      spService.getTaxonomy(OrganTaxonomyIds.BusinessArea),
      spService.getItems<ILookupEXTERNALField>(`${this.props.context.pageContext.site.serverRelativeUrl}${EXTERNALList.ListStatusUrl}`, EXTERNALList.getAll()),
      bkService.getAreaRelationsInfo('Contoso'),
      spService.getItems<ILookupEXTERNALField>(`${this.props.context.pageContext.site.serverRelativeUrl}${EXTERNALList.ListTipoDocumentUrl}`,EXTERNALList.getAll())
    ]);

    const organsTermInfo: Term[] = organsTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    //const organsTypeTermInfo: Term[] = organsTypeTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    //const organsMinistryTermInfo: Term[] = ministryTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const usersRoleTermInfo: Term[] = roleTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const BusinessAreaTermInfo: Term[] = areaSecTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    //const sitacionTermInfo: Term[] = StatusTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));

    this._organListDictionary = organsTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this._organTypeListDictionary = organsTypeLookup.reduce((acc, item) => {
      acc[item.Id] = item.descripcion;
      return acc;
    }, {} as Record<string, string>);

    this._ministryDictionary = ministryLookup.reduce((acc, item) => {
      acc[item.Id] = item.descripcion;
      return acc;
    }, {} as Record<string, string>);

    this._documentTypesDictionary = documentTypes.reduce((acc, item) => {
      acc[item.Id] = item.descripcion;
      return acc;
    }, {} as Record<string, string>);


    const currentAreaMinistryRelatio = Object.values(
      BusinessAreaRelation.reduce((acc, item) => {
        const current = acc[item.BusinessArea];
        if (!current || (new Date(item.StartDate) > new Date(current.StartDate))) {
          acc[item.BusinessArea] = item;
        }
        return acc;
      }, {} as { [key: string]: IRelationBusinessArea })
    );
    this._BusinessAreaDivisionRelations = currentAreaMinistryRelatio;

    this.setState({
      organTypeFilterOptions: organsTypeLookup,
      ministryFilterOptions: ministryLookup,
      organFilterOptions: organsTermInfo,
      roleFilterOptions: usersRoleTermInfo,
      BusinessAreaOptions: BusinessAreaTermInfo,
      StatusFilterOptions: StatusTerms,
    });
  }


  private onResetFilters(): void {
    const { organBackList } = this.state;
      this.setState({
        searchText: "",
        organBackResultList: [...organBackList],
        searchFilters: Certification.eCNTyFilters,
        enableRemoveFilters: false,
        currentPage: 1
      });
  }

  private onSearch(searchText: string): void {
    this.setState({ searchText: searchText },
      () => this.onFilter(this.state.searchFilters)
    )
  }

  private escapeRegExp(txt:string):string {
    return txt.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  private onFilter(filters: ISearchFilters): void {
    const { searchText, organBackList } = this.state;
    const { OrganTypeEXTERNAL, BusinessArea, CodigoDepartment, FechaConstitucionDesde,FechaConstitucionHasta,FechaExtincionDesde,FechaExtincionHasta,FechaInscripcionDesde,FechaInscripcionHasta } = filters;
    let filteredOrgans = [...organBackList];
    if (searchText) {
      const escapedSearchText = this.escapeRegExp(searchText);
      const regex = new RegExp(escapedSearchText, "i");
      filteredOrgans = [...organBackList].filter(organ => regex.test(this._organListDictionary[organ.Department]))
    }
    /*
    if (OrganType && OrganType.length > 0) {
      filteredOrgans = filteredOrgans.filter(item =>
        OrganType.some(m => m === item.TipoDepartment)
      );
    }
    if (Ministry && Ministry.length > 0) {
      filteredOrgans = filteredOrgans.filter(item =>
        Ministry.some(m => m === item.Division)
      );
    }*/
    if (OrganTypeEXTERNAL && OrganTypeEXTERNAL.length > 0) {
      filteredOrgans = filteredOrgans.filter(item =>
        OrganTypeEXTERNAL.some(m => m === item.TipoDepartmentEXTERNAL)
      );
    }
    if (BusinessArea && BusinessArea.length > 0) {
      filteredOrgans = filteredOrgans.filter(item =>
        BusinessArea.some(m => m === item.BusinessArea)
      );
    }/*
    if (User) {
      filteredOrgans = filteredOrgans.filter(item =>
        item?.Users?.find(u => u.UserPrincipalName === User)
      )
    }*/
    if(CodigoDepartment){
      const escapedSearchText = this.escapeRegExp(CodigoDepartment);
      const regex = new RegExp(escapedSearchText, "i");
      filteredOrgans = filteredOrgans.filter(organ => regex.test(organ.CodigoDepartment))
    }
    if(FechaConstitucionDesde){
      filteredOrgans = filteredOrgans.filter(item => item.FechaConstitucion && new Date(item.FechaConstitucion) >= new Date(FechaConstitucionDesde));
    }

    if(FechaConstitucionHasta){
      filteredOrgans = filteredOrgans.filter(item => item.FechaConstitucion && new Date(item.FechaConstitucion) <= new Date(FechaConstitucionHasta));
    }

    if(FechaExtincionDesde){
      filteredOrgans = filteredOrgans.filter(item => item.FechaExtincion && new Date(item.FechaExtincion) >= new Date(FechaExtincionDesde));
    }

    if(FechaExtincionHasta){
      filteredOrgans = filteredOrgans.filter(item => item.FechaExtincion && new Date(item.FechaExtincion) <= new Date(FechaExtincionHasta));
    }

    if(FechaInscripcionDesde){
      filteredOrgans = filteredOrgans.filter(item => item.FechaInscripcion && new Date(item.FechaInscripcion) >= new Date(FechaInscripcionDesde));
    }

    if(FechaInscripcionHasta){
      filteredOrgans = filteredOrgans.filter(item => item.FechaInscripcion && new Date(item.FechaInscripcion) <= new Date(FechaInscripcionHasta));
    }

    this.setState({ organBackResultList: [...filteredOrgans], searchFilters: filters, currentPage: 1 });
    this.checkFilterAndSearch(filters);
  }

  private async onDownloadBodyCertificate(bodyId: string, isEXTERNAL?: boolean): promise<ArrayBuffer> {
    return this.props.bkService.downloadBodyCertificate('Contoso', bodyId, isEXTERNAL);
  }

  private async onManageItem(itemId: string, action: string, item?: IOrganEXTERNALResult): promise<void> {
    const { bkService } = this.props;
    if (action === "Open") {
      this.setState({
        dialogIsOpen: true,
        managedItem: {
          itemId,
          organDialogContent:  this.getCompleteOrgan(itemId),
          isNewItem: false
        },
        errorUpdating: false
      });
    }
    else if (action === "OpenNew") {
      this.setState({
        dialogIsOpen: true,
        managedItem: {
          itemId: "",
          organDialogContent: 
            {
              CodigoDepartment: '',
              BodyId: '',
              Department: '',
              TipoDepartmentEXTERNAL: '',
              SecretariaText: '',
              Division: '',
              DocumentSetDescription: '',
              Intersectorial: false,
              dir: '',
              SIA: '',
              Abreviatura: '',
              Activo: false,
              Observaciones: '',
              IdCertificateBodies: 0,
              BusinessArea: '',
              NombreDepartment: '',
              StatusDepartment: '',
              Inscrito:false,
            },
          isNewItem: true
        }
      });
    }
    else if (action === "Close") {
      this.setState({
        dialogIsOpen: false,
        managedItem: undefined
      });
    }
    else if (action === "Updateproperties") {
      if (item) {
        this.setState({
          isLoadingUpdates: true,
          searchFilters: Certification.eCNTyFilters,
          errorUpdating: false
        });
        bkService.updateEXTERNALOrganInfo(item, 'Contoso').then(() => {
          this.setState({ dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, currentPage: 1 });
          this.loadData();
        }
        ).catch(ex => {
          console.log(ex);
          this.setState({ isLoadingUpdates: false, errorUpdating: true });
        })
      }
    }
    else if (action === "CreateNew") {
      if (item) {
        this.setState({
          isLoadingUpdates: true,
          searchFilters: Certification.eCNTyFilters,
          errorUpdating: false
        });

        bkService.createEXTERNALOrgan(item, 'Contoso').then(() => {
          this.setState({ dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, currentPage: 1 });
          //this.loadData();
          window.location.reload();
        }
        ).catch(ex => {
          console.log(ex);
          this.setState({ isLoadingUpdates: false, errorUpdating: true });
        })
      }
    }
  }

  private getCompleteOrgan(itemId: string): IOrganEXTERNALResult | undefined {
    const currentOrgan = this.state.organBackList.find(org => org.CodigoDepartment === itemId);
    if (currentOrgan) {
      console.log(currentOrgan);
      return (currentOrgan);
    } else {
      return undefined;
    }
  }

  private checkFilterAndSearch(filters: ISearchFilters): void {
    const { searchText } = this.state;

    if (searchText.trim() !== "" || !_.isEqual(filters, Certification.eCNTyFilters)) {
      this.setState({ enableRemoveFilters: true });
    } else {
      this.setState({ enableRemoveFilters: false });
    }
  }

  private setPaging(page: number): void {
    this.setState({ currentPage: page });
  }


  public render(): React.ReactElement<ICertificationAppprops> {
    const { componentTitle, isMobile, itemsPerPage } = this.props;
    const { organBackList, StatusFilterOptions, errorUpdating, showLoadDataError, loadDataError, isLoading, currentPage, enableRemoveFilters, searchFilters, BusinessAreaOptions, searchText, managedItem, dialogIsOpen, organFilterOptions, ministryFilterOptions, organTypeFilterOptions, documentDialog } = this.state;
    return (
      <div className={styles.certificationApp} >
        {organBackList && organBackList.length > 0 && documentDialog &&
          <EXTERNALDocumentForm 
            isOpen={true} 
            organs={organBackList} 
            onClose={() => 
              this.setState({documentDialog: false})
             }
            organTypes={this._organTypeListDictionary}
            documentTypes={this._documentTypesDictionary}
            spService={this.props.spService}
          ></EXTERNALDocumentForm>
        }
        <div className={styles.wpTitle}>
          {
          (componentTitle && componentTitle !== "") ?
            <div>{this.props.componentTitle}</div>
            :
            <div></div>
          }
          {!isLoading && !showLoadDataError &&
            <div className={styles.newElementsButtons}>
              <Button appearance={"secondary"} icon={<DocumentAddRegular/>} onClick={() => this.setState({documentDialog:true})} style={{marginRight:"20px"}}>{strings.UploadDocument}</Button>
              <Button appearance={"primary"} icon={<PersonAddRegular />} onClick={() => this.onManageItem("", "OpenNew")}>{strings.NewOrgan}</Button>
            </div>
          } 
        </div>
        {(isLoading || showLoadDataError) ?
          <div>
            {showLoadDataError ?
              <div className={styles.errorLabel}>
                {strings.LoadingError}
                {loadDataError && <test style={{ whiteSpace: 'test-wrap', margin: '8px 0 0', fontSize: '12px', fontFamily: 'monospace', wordBreak: 'break-all' }}>{loadDataError}</test>}
              </div>
              :
              <Spinner labelPosition="below" label={strings.LoadingLabel} />
            }

          </div>
          :
          <>
            <div className={styles.newElementsButtons} hidden>
              <Button icon={<AddRegular />}
                onClick={() => this.onManageItem("", "OpenNew")}
              >
                {strings.NewOrgan}
              </Button>
            </div>
            <SearchBar
              onSearch={this.onSearch.bind(this)}
              enableRemoveFilters={enableRemoveFilters}
              isMobile={isMobile}
              onResetFilters={this.onResetFilters.bind(this)}
            />
            <Divider className={"styles.dividerHorizontal"} />
            <FilterZone
              organTypeFilterOptions={organTypeFilterOptions}
              BusinessAreaOptions={BusinessAreaOptions}
              onApplyFilters={this.onFilter.bind(this)}
              searchText={searchText}
              searchFilters={searchFilters}
            />
            <Divider className={"styles.dividerHorizontal"} />
            <ResultZone
              onManageItem={this.onManageItem.bind(this)}
              isMobile={isMobile}
              numberPageItems={itemsPerPage}
              currentPage={currentPage}
              onUpdatePage={this.setPaging.bind(this)}
              organBackList={this.state.organBackResultList}
              organListDictionary={this._organListDictionary}
              organTypeListDictionary={this._organTypeListDictionary}
              ministryDictionary={this._ministryDictionary}
              onDownloadCertificate={this.onDownloadBodyCertificate.bind(this)}
            />
            {
              <Dialog open={dialogIsOpen}>
                <DialogSurface className={styles.managementDialog}>
                  <DialogBody>
                    <DialogTitle>
                      {strings.ManageOrgan}
                    </DialogTitle>
                    <EXTERNALOrganDetail
                      managedItem={managedItem}
                      onManageItem={this.onManageItem.bind(this)}
                      organTypeFilterOptions={organTypeFilterOptions}
                      ministryFilterOptions={ministryFilterOptions}
                      organFilterOptions={organFilterOptions}
                      BusinessAreaOptions={BusinessAreaOptions}
                      StatusOptions={StatusFilterOptions}
                      isMobile={isMobile}
                      organDictionary={this._organListDictionary}
                      organTypeDictionary={this._organTypeListDictionary}
                      ministryDictionary={this._ministryDictionary}
                      errorUpdating={errorUpdating}
                      allOrganCodes={this._allOrganCodes}
                      allOrgansData={this.state.organBackList}
                      BusinessAreaMiniesterioRelations={this._BusinessAreaDivisionRelations}
                    />
                  </DialogBody>
                </DialogSurface>
              </Dialog >
            }
          </>

        }
      </div >
    );
  }
}
