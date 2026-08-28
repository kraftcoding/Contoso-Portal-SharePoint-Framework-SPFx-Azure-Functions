import * as React from 'react';
import styles from './AdministrationApp.module.scss';
import {
  IAdministrationAppprops,
  IAdministrationAppState
} from './IAdministrationApp';
import SearchBar from './SearchBar/SearchBar';
import FilterZone from './FilterZone/FilterZone';
import ResultZone from './ResultZone/ResultZone';
import { Administration, componentTypeMode, ISearchFilters, SearchType, OrganTaxonomyIds, LicensesTypes, IUserResult, RoleTaxonomyIds, IOrganResult, IUsersOrganRelation, IUserPermissionsCheck, IPermissionCheckResult, BodyUserInfo, ILicenseInfo } from '../models/AdministrationAppModels';
import { Button, Dialog, DialogBody, DialogSurface, DialogTitle, Divider, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner } from '@fluentui/react-components';
import _ from 'lodash';
import DepartmentDetails from '../../../components/DepartmentDetails/DepartmentDetails';
import strings from 'AdministrationAppWebPartStrings';
import { getBodyIdFormUrl } from '../../../service/RoleService';
import { Term } from '../../../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { PeopleAddRegular, PeopleCommunityAddRegular, PersonAddRegular } from '@fluentui/react-icons';
import NewUserForm from '../../../components/NewUserForm/NewUserForm';
import NewUsersFromExcel from '../../../components/NewUsersFromExcel/NewUsersFromExcel';
import { serializeError } from '../../../utils/errorUtils';

export default class AdministrationApp extends React.Component<IAdministrationAppprops, IAdministrationAppState> {

  private _organListDictionary: Record<string, string>;
  private _organTypeListDictionary: Record<string, string>;
  private _ministryDictionary: Record<string, string>;
  private _communityListDictionary: Record<string, string>;
  private _roleListDictionary: Record<string, string>;

  private _usersListDictionary: Record<string, IUserResult>;
  private _invitationMessage: string;
  private _organListWithBodyId: Term[];
  constructor(props: IAdministrationAppprops) {
    super(props);
    this.state = {
      searchType: this.props.componentMode === componentTypeMode.All ? SearchType.Organ : this.props.componentMode === componentTypeMode.Organs ? SearchType.Organ : SearchType.User,
      searchText: "",
      organsList: [],
      usersList: [],
      organTypeFilterOptions: [],
      ministryFilterOptions: [],
      organFilterOptions: [],
      ccaaFilterOptions: [],
      roleFilterOptions: [],
      BusinessAreaOptions: [],
      resultsOrgansList: [],
      resultsUsersList: [],
      searchFilters: Administration.eCNTyFilters,
      enableRemoveFilters: false,
      dialogIsOpen: false,
      managedItem: {
        itemId: "",
        itemType: SearchType.None,
        organDialogContent: undefined,
        userDialogContent: undefined
      },
      currentPage: 1,
      secretariaOptions: [],
      legistatureOptions: [],
      isLoading: true,
      organBackList: [],
      userBackList: [],
      organBackResultList: [],
      userBackResultList: [],
      showLoadDataError: false,
      loadDataError: null,
      isLoadingUpdates: true,
      errorUpdating: false,
      availableLicenses: [],
      newUserDialogIsOpen: false,
      errorCreatingUser: false,
      newUsersFromExcel: false
    };
  }

  async componentDidMount(): promise<void> {
    await this.getTaxonomies();
    void this.loadData(); 
    this.props.bkService.getUserInvitationMessage("/sites/Contoso/NotificationUserConfig/Messages/InvitacionUsuarios.txt").then(value => {
      this._invitationMessage = value
    }).catch(ex => console.log(ex))
  }

  private loadData(): void {
    const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
    const request: promise<any>[] = [];
    if (this.props.showEXTERNAL) {
      request.push(this.props.bkService.getEXTERNALOrgansInfo("Contoso"));
    } else {
      request.push(
        this.props.bkService.getOrganInfo(bodyId ? bodyId : "Contoso"),
        this.props.bkService.getUsersInfo(bodyId ? bodyId : "Contoso"),
        this.props.bkService.getAvailableLicenses(bodyId ? bodyId : "Contoso")
      )
    }

    promise.all(request).then(results => {
      const organs: IOrganResult[] = [];
      let users: Record<string, IUserResult> = {};
      let licenseInfo: ILicenseInfo[]=[];
      if (this.props.showEXTERNAL) {
        organs.push(...results[0]);
      } else {
        organs.push(...results[0]);
        users = results[1];
        licenseInfo = results[2];
      }
      const orderedOrgans = organs.sort((a, b) => this._organListDictionary[a.Department].localeCompare(this._organListDictionary[b.Department]));
      console.log(orderedOrgans);
      orderedOrgans.forEach(org => {
        org.Users.forEach(u => {
          if (u.UserPrincipalName && users[u.UserPrincipalName.toLowerCase()]) {
            const userRole = this.onMaprodleFromBack(u.UserRoles);
            if (userRole) {
              u.RoleId = userRole.roleId;
              u.Ccaa = users[u.UserPrincipalName.toLowerCase()].Office;
            }
            users[u.UserPrincipalName.toLowerCase()].Organ
              ?
              users[u.UserPrincipalName.toLowerCase()].Organ?.push({ OrganId: org.Department, DepartmentName: this._organListDictionary[org.Department], Role: userRole ? userRole.role : "", RoleId: userRole ? userRole.roleId : "" })
              :
              users[u.UserPrincipalName.toLowerCase()].Organ = [{ OrganId: org.Department, DepartmentName: this._organListDictionary[org.Department], Role: userRole ? userRole.role : "", RoleId: userRole ? userRole.roleId : "" }]
          }
        })
        org.Users = org.Users.filter(user => users[user.UserPrincipalName.toLowerCase()]).sort((a, b) => users[a.UserPrincipalName.toLowerCase()]?.DisplayName.localeCompare(users[b.UserPrincipalName.toLowerCase()]?.DisplayName));
        org.NombreDepartment = this._organListDictionary[org.Department];
      });
      this._usersListDictionary = users;
      const dicToArrayUsers = Object.values(users);
      const orderedUsers = dicToArrayUsers.sort((a, b) => a.DisplayName.localeCompare(b.DisplayName));

      const organsFilterFromList:Term[] = orderedOrgans.map((org:IOrganResult) => ({key:org.Department, text:org.NombreDepartment}));
      this._organListWithBodyId = orderedOrgans.map((org:IOrganResult) => ({key:org.BodyId, text:org.NombreDepartment}));
      //organFilterOptions: organsTermInfo,
      this.setState({organFilterOptions: organsFilterFromList, organBackList: orderedOrgans, organBackResultList: orderedOrgans, userBackList: orderedUsers, userBackResultList: orderedUsers, availableLicenses:licenseInfo, isLoading: false });
    }).catch(ex => {
      this.setState({ showLoadDataError: true, isLoading: false, loadDataError: serializeError(ex) });
      console.error(ex);
    });
  }

  private async getTaxonomies(): promise<void> {
    const { spService } = this.props;

    const [organsTerms, organsTypeTerms, ministryTerms, secretariaTerms, PeriodTerms, communityTerms, roleTerms, areaSecTerms]: [ITermInfo[], ITermInfo[], ITermInfo[], ITermInfo[], ITermInfo[], ITermInfo[], ITermInfo[], ITermInfo[]] = await promise.all([
      spService.getTaxonomy(OrganTaxonomyIds.Organ),
      spService.getTaxonomy(OrganTaxonomyIds.OrganType),
      spService.getTaxonomy(OrganTaxonomyIds.Ministry),
      spService.getTaxonomy(OrganTaxonomyIds.Secretaria),
      spService.getTaxonomy(OrganTaxonomyIds.Period),
      spService.getTaxonomy(OrganTaxonomyIds.Community),
      spService.getTaxonomy(OrganTaxonomyIds.Roles),
      spService.getTaxonomy(OrganTaxonomyIds.BusinessArea),
    ]);

    const organsTermInfo: Term[] = organsTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const organsTypeTermInfo: Term[] = organsTypeTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const organsMinistryTermInfo: Term[] = ministryTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const organsSecretariaTermInfo: Term[] = secretariaTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const organsPeriodTermInfo: Term[] = PeriodTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const usersCommunityTermInfo: Term[] = communityTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const usersRoleTermInfo: Term[] = roleTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));
    const BusinessAreaTermInfo: Term[] = areaSecTerms?.map((tag: ITermInfo) => Term.mapSPToTerm(tag, this.props.context.pageContext.cultureInfo.currentUICultureName));

    this._organListDictionary = organsTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this._organTypeListDictionary = organsTypeTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this._ministryDictionary = organsMinistryTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this._communityListDictionary = usersCommunityTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this._roleListDictionary = usersRoleTermInfo.reduce((acc, item) => {
      acc[item.key] = item.text;
      return acc;
    }, {} as Record<string, string>);

    this.setState({
      organTypeFilterOptions: organsTypeTermInfo,
      ministryFilterOptions: organsMinistryTermInfo,
      //organFilterOptions: organsTermInfo,
      ccaaFilterOptions: usersCommunityTermInfo,
      roleFilterOptions: usersRoleTermInfo,
      secretariaOptions: organsSecretariaTermInfo,
      legistatureOptions: organsPeriodTermInfo,
      BusinessAreaOptions: BusinessAreaTermInfo
    });
  }

  private onMaprodleFromBack(roles: string[]): ({ roleId: string, role: string } | null) {
    if (roles.indexOf('schedulers') > -1) {
      return { role: this._roleListDictionary[RoleTaxonomyIds.Convocante], roleId: RoleTaxonomyIds.Convocante }
    } else if (roles.indexOf('gestorschedulers') > -1) {
      return { role: this._roleListDictionary[RoleTaxonomyIds.GestorConvocante], roleId: RoleTaxonomyIds.GestorConvocante }
    } else if (roles.indexOf('members') > -1) {
      return { role: this._roleListDictionary[RoleTaxonomyIds.members], roleId: RoleTaxonomyIds.members }
    } else if (roles.indexOf('asistentemembers') > -1) {
      return { role: this._roleListDictionary[RoleTaxonomyIds.AsistenteMiembro], roleId: RoleTaxonomyIds.AsistenteMiembro }
    } else if (roles.indexOf('guests') > -1) {
      return { role: this._roleListDictionary[RoleTaxonomyIds.Invitado], roleId: RoleTaxonomyIds.Invitado }
    } else {
      return null
    }
  }

  private onMaprodleTaxIdToBack(roleTaxId: string): string {
    if (roleTaxId === RoleTaxonomyIds.Convocante) {
      return "schedulers";
    } else if (roleTaxId === RoleTaxonomyIds.GestorConvocante) {
      return "gestorschedulers";
    } else if (roleTaxId === RoleTaxonomyIds.members) {
      return "members";
    } else if (roleTaxId === RoleTaxonomyIds.AsistenteMiembro) {
      return "asistentemembers";
    } else if (roleTaxId === RoleTaxonomyIds.Invitado) {
      return "guests";
    } else {
      return roleTaxId;
    }
  }

  private onSelectSearchType(searchType: SearchType): void {
    this.setState({
      searchType,
      searchText: "",
      organBackResultList: this.state.organBackList,
      userBackResultList: this.state.userBackList,
      searchFilters: Administration.eCNTyFilters,
      currentPage: 1,
      managedItem: undefined
    });
  }

  private onResetFilters(): void {
    const { searchType, userBackList, organBackList } = this.state;
    if (searchType === SearchType.Organ) {
      this.setState({
        searchText: "",
        organBackResultList: [...organBackList],
        searchFilters: Administration.eCNTyFilters,
        enableRemoveFilters: false,
        currentPage: 1
      });
    } else if (searchType === SearchType.User) {
      this.setState({
        searchText: "",
        userBackResultList: [...userBackList],
        searchFilters: Administration.eCNTyFilters,
        enableRemoveFilters: false,
        currentPage: 1
      });
    }
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
    const { searchText, organBackList, userBackList } = this.state;
    const { OrganType, AutonomousCommunity, Rol, Organ, Ministry, BusinessArea, User, License } = filters;
    console.log(filters);
    if (this.state.searchType === SearchType.Organ) {
      let filteredOrgans = [...organBackList];
      if (searchText) {
        const escapedSearchText = this.escapeRegExp(searchText);
        const regex = new RegExp(escapedSearchText, "i");

        filteredOrgans = [...organBackList].filter(organ => regex.test(this._organListDictionary[organ.Department]))
      }
      if (OrganType && OrganType.length > 0) {
        filteredOrgans = filteredOrgans.filter(item =>
          OrganType.some(m => m === item.TipoDepartment)
        );
      }
      if (Ministry && Ministry.length > 0) {
        filteredOrgans = filteredOrgans.filter(item =>
          Ministry.some(m => m === item.Division)
        );
      }
      if (BusinessArea && BusinessArea.length > 0) {
        filteredOrgans = filteredOrgans.filter(item =>
          BusinessArea.some(m => m === item.BusinessArea)
        );
      }
      if (User) {
        filteredOrgans = filteredOrgans.filter(item =>
          item?.Users?.find(u => u.UserPrincipalName === User)
        )
      }
      this.setState({ organBackResultList: [...filteredOrgans], searchFilters: filters, currentPage: 1 });
    } else {
      let filteredUsers = [...userBackList];
      if (searchText) {
        const escapedSearchText = this.escapeRegExp(searchText);
        const regex = new RegExp(escapedSearchText, "i");
        filteredUsers = [...userBackList].filter(user => regex.test(user.DisplayName))
      }

      if (Organ && Organ.length > 0) {
        filteredUsers = filteredUsers.filter(item =>
          Organ?.some(u => item?.Organ?.find(or => u === or.OrganId))
        )
      }
      if (AutonomousCommunity && AutonomousCommunity.length > 0) {
        filteredUsers = filteredUsers.filter(item =>
          AutonomousCommunity.some(u => u === item.Office)
        );
      }

      if (Rol && Rol.length > 0) {
        filteredUsers = filteredUsers.filter(item =>
          Rol.some(u => item?.Organ?.find(or => u === or.RoleId))
        );
      }
      if (License && License.length > 0) {
        if (License.some(el => el ===LicensesTypes.E5)) {
          if(License.some(el => el ===LicensesTypes.E3)){
            filteredUsers = filteredUsers.filter(item => (item.AssignedLicenses.includes(LicensesTypes.E5) || item.AssignedLicenses.includes(LicensesTypes.E5Developer) || item.AssignedLicenses.includes(LicensesTypes.E3)));
          }else{
            filteredUsers = filteredUsers.filter(item => (item.AssignedLicenses.includes(LicensesTypes.E5) || item.AssignedLicenses.includes(LicensesTypes.E5Developer)));
          }
        } else if(License.some(el => el ===LicensesTypes.E3)){
          filteredUsers = filteredUsers.filter(item => (item.AssignedLicenses.includes(LicensesTypes.E3)));
        }
      }
      this.setState({ userBackResultList: [...filteredUsers], searchFilters: filters, currentPage: 1 });
    }
    this.checkFilterAndSearch(filters);
  }
  private async onCheckPermissions(item: IUserPermissionsCheck): promise<IPermissionCheckResult> {
    let bodyId = item.BodyId;
    if (this.state.searchType === SearchType.User && item.BodyId) {
      const getBodyId = this.state.organBackList.find(or => or.Department === item.BodyId)?.BodyId;
      if (getBodyId) {
        bodyId = getBodyId;
      }
    }
    const role = item.Role ? this.onMaprodleTaxIdToBack(item.Role) : "";
    const itemToCheck: IUsersOrganRelation = {
      UserUpn: item?.UserPrincipalName ? item.UserPrincipalName : "",
      BodyId: bodyId ? bodyId : "",
      Role: role
    }
    return this.props.bkService.checkUserHasPermissions(itemToCheck);
  }

  private async onUpdateLicense(upn: string, license: string): promise<void>{
    this.setState({
      isLoadingUpdates: true,
      errorUpdating: false
    });
    const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
    this.props.bkService.manageUserLicense(bodyId ? bodyId : 'Contoso', upn, license).then(res => {
      this.setState({searchFilters: Administration.eCNTyFilters, dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: SearchType.User, currentPage: 1 });
      this.loadData();
    }
    ).catch(ex => {
      console.log(ex);
      this.setState({ isLoadingUpdates: false, errorUpdating: true });
    })
  }

  private async onDownloadBodyCertificate(bodyId: string, isEXTERNAL?: boolean): promise<ArrayBuffer> {
    return this.props.bkService.downloadBodyCertificate('Contoso', bodyId, isEXTERNAL);
  }

  private async onManageItem(itemId: string, itemType: SearchType, action: string, item?: IOrganResult | IUserResult | IUsersOrganRelation[]): promise<void> {
    const { bkService } = this.props;
    if (action === "Open") {
      this.setState({
        dialogIsOpen: true,
        managedItem: {
          itemId,
          itemType,
          organDialogContent: itemType === SearchType.Organ ? this.getCompleteOrgan(itemId) : undefined,
          userDialogContent: itemType === SearchType.User ? this.getCompleteUser(itemId) : undefined,
        },
        errorUpdating: false
      });
    }
    else if (action === "OpenNew") {
      this.setState({
        newUserDialogIsOpen: true,
        errorCreatingUser: false,
        newUsersFromExcel: false
      });
    }
    else if (action === "OpenNewMultiple"){
      this.setState({
        newUserDialogIsOpen: true,
        errorCreatingUser: false,
        newUsersFromExcel: true
      });
    }
    else if (action === "Close") {
      this.setState({
        dialogIsOpen: false,
        managedItem: undefined,
        newUserDialogIsOpen: false,
        errorCreatingUser: false
      });
    }
    else if (action === "Updateproperties") {
      console.log(item);
      if (item) {
        if (itemType === SearchType.Organ) {
          const updateItem: IOrganResult = item as IOrganResult;
          this.setState({
            isLoadingUpdates: true,
            searchFilters: Administration.eCNTyFilters,
            errorUpdating: false
          });

          if (!updateItem.IsEXTERNAL) {
            bkService.updateOrganInfo(updateItem, 'Contoso').then(() => {
              this.setState({ dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: SearchType.Organ, currentPage: 1 });
              this.loadData();
            }
            ).catch(ex => {
              console.log(ex);
              this.setState({ isLoadingUpdates: false, errorUpdating: true });
            })
          }

          console.log(updateItem);

        } else if (itemType === SearchType.User) {
          this.setState({
            isLoadingUpdates: true,
            searchFilters: Administration.eCNTyFilters,
            errorUpdating: false
          });
          const updateItem: IUserResult = item as IUserResult;
          bkService.updateUserprofile(updateItem, 'Contoso').then(() => {
            this.setState({ dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: SearchType.User, currentPage: 1 });
            this.loadData();
          }
          ).catch(ex => {
            console.log(ex);
            this.setState({ isLoadingUpdates: false, errorUpdating: true });
          })
        }
      }
    } else if (action === "UpdateMember") {
      if (item) {
        const memberOrgans: IUsersOrganRelation[] = item as IUsersOrganRelation[];
        const organUserRelation: IUsersOrganRelation[] = memberOrgans.map(e => {
          let currentItem = e;
          if (e.Role) {
            currentItem = { ...currentItem, Role: this.onMaprodleTaxIdToBack(e.Role) };
          } else {
            currentItem = { ...currentItem, Role: "" };
          }
          if (itemType === SearchType.User) {
            const organBodyId = this.state.organBackList.find(or => or.Department === e.BodyId)?.BodyId;
            if (organBodyId) {
              currentItem = { ...currentItem, BodyId: organBodyId };
            }
          }
          return currentItem;
        });

        console.log(organUserRelation);
        bkService.updateUsersInOrgan(organUserRelation, 'Contoso').then(() => {
          this.setState({searchFilters: Administration.eCNTyFilters, dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: this.state.searchType, currentPage: 1 });
          this.loadData();
        }
        ).catch(ex => {
          console.log(ex);
          this.setState({ isLoadingUpdates: false, errorUpdating: true});
        })
      }
    }
  }

  private getCompleteOrgan(itemId: string): IOrganResult | undefined {
    const currentOrgan = this.state.organBackList.find(org => org.BodyId === itemId);
    if (currentOrgan) {
      console.log(currentOrgan);
      return (currentOrgan);
    } else {
      return undefined;
    }
  }

  private getCompleteUser(itemId: string): IUserResult | undefined {
    // TODO: Hacer una llamada al backend para obtener el usuario
    const currentUser = this.state.userBackList.find(usr => usr.UserPrincipalName === itemId);
    if (currentUser) {
      return currentUser
    } else {
      return undefined
    }
  }

  private onCloseNewUserPanel():void{
    this.setState({
      dialogIsOpen: false,
      managedItem: undefined,
      newUserDialogIsOpen: false
    })
  }

  private onReloadItems(): void{
    this.setState({ dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: this.state.searchType, currentPage: 1, newUserDialogIsOpen: false });
    this.loadData();
  }

  private onSaveNewUserPanel(user:BodyUserInfo):void{
    console.log(user);
    const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
    //createUser(item:BodyUserInformation, bodyId: string):promise<void>;
    this.setState({ errorCreatingUser: false });
    let userToCreate = {...user};
    if(userToCreate.UserBodyRole && userToCreate.UserBodyRole.BodyId && userToCreate.UserBodyRole.Role){
      const bodyIdToAdd =  this.state.organBackList.find(or => or.Department === userToCreate.UserBodyRole?.BodyId)?.BodyId;
      const roleToAdd = this.onMaprodleTaxIdToBack(userToCreate.UserBodyRole.Role)
      userToCreate.UserBodyRole.BodyId = bodyIdToAdd;
      userToCreate.UserBodyRole.Role = roleToAdd;
    }
    this.props.bkService.createUser(userToCreate, bodyId ? bodyId : 'Contoso').then(() => 
      {
        this.setState({searchFilters: Administration.eCNTyFilters, dialogIsOpen: false, managedItem: undefined, isLoading: true, isLoadingUpdates: false, searchType: this.state.searchType, currentPage: 1, newUserDialogIsOpen: false });
        this.loadData();
      }
    ).catch(ex => {
      console.log(ex);
      this.setState({ isLoadingUpdates: false, errorCreatingUser: true });
    })
  }

  private checkFilterAndSearch(filters: ISearchFilters): void {
    const { searchText } = this.state;

    if (searchText.trim() !== "" || !_.isEqual(filters, Administration.eCNTyFilters)) {
      this.setState({ enableRemoveFilters: true });
    } else {
      this.setState({ enableRemoveFilters: false });
    }
  }

  private setPaging(page: number): void {
    this.setState({ currentPage: page });
  }

  public render(): React.ReactElement<IAdministrationAppprops> {
    const { componentTitle, componentMode, isMobile, itemsPerPage, showEXTERNAL } = this.props;
    const {availableLicenses, newUsersFromExcel, newUserDialogIsOpen, errorCreatingUser, errorUpdating, showLoadDataError, loadDataError, isLoading, legistatureOptions, secretariaOptions, currentPage, enableRemoveFilters, searchFilters, BusinessAreaOptions, searchType, searchText, managedItem, dialogIsOpen, ccaaFilterOptions, roleFilterOptions, organFilterOptions, ministryFilterOptions, organTypeFilterOptions } = this.state;
    console.log(errorCreatingUser);
    return (
      <div className={styles.administrationApp} >
        
          <div className={styles.wpTitle}>
            {
            (componentTitle && componentTitle !== "") ?
              <div>{this.props.componentTitle}</div>
              :
              <div></div>
            }
            {!isLoading && !showLoadDataError &&
              <div className={styles.newElementsButtons} hidden={searchType !== SearchType.User}>
                {isMobile ? 
                  <>
                  <Button appearance={"primary"} icon={<PersonAddRegular />} onClick={() => this.onManageItem("", searchType, "OpenNew")}>{strings.NewUser}</Button>
                  <Button appearance={"primary"} icon={<PeopleCommunityAddRegular />} onClick={() => this.onManageItem("", searchType, "OpenNewMultiple")}>{strings.NewMultipleUsers}</Button>
                  </>
                :
                  <Menu>
                    <MenuTrigger>
                        <Button appearance={"primary"} icon={<PeopleAddRegular />}>{strings.NewUsers}</Button>
                    </MenuTrigger>
                    <MenuPopover>
                      <MenuList>
                        <MenuItem icon={<PersonAddRegular />} onClick={() => this.onManageItem("", searchType, "OpenNew")}>{strings.NewUser}</MenuItem>
                        <MenuItem icon={<PeopleCommunityAddRegular />} onClick={() => this.onManageItem("", searchType, "OpenNewMultiple")}>{strings.NewMultipleUsers}</MenuItem>
                      </MenuList>
                    </MenuPopover>
                  </Menu>
                }
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
            <SearchBar
              onSelectSearchType={this.onSelectSearchType.bind(this)}
              onSearch={this.onSearch.bind(this)}
              showTypeSelector={componentMode === componentTypeMode.All}
              enableRemoveFilters={enableRemoveFilters}
              isMobile={isMobile}
              searchType={searchType}
              onResetFilters={this.onResetFilters.bind(this)}
            />
            <Divider className={"styles.dividerHorizontal"} />
            <FilterZone
              searchType={searchType}
              organTypeFilterOptions={organTypeFilterOptions}
              BusinessAreaOptions={BusinessAreaOptions}
              organFilterOptions={organFilterOptions}
              ccaaFilterOptions={ccaaFilterOptions}
              roleFilterOptions={roleFilterOptions}
              onApplyFilters={this.onFilter.bind(this)}
              searchText={searchText}
              searchFilters={searchFilters}
              showEXTERNAL={showEXTERNAL}
            />
            <Divider className={"styles.dividerHorizontal"} />
            {
              searchType !== SearchType.None &&
              <ResultZone
                searchType={searchType}
                onManageItem={this.onManageItem.bind(this)}
                isMobile={isMobile}
                numberPageItems={itemsPerPage}
                currentPage={currentPage}
                onUpdatePage={this.setPaging.bind(this)}
                userBackList={this.state.userBackResultList}
                organBackList={this.state.organBackResultList}
                organListDictionary={this._organListDictionary}
                organTypeListDictionary={this._organTypeListDictionary}
                ministryDictionary={this._ministryDictionary}
                communityListDictionary={this._communityListDictionary}
                roleListDictionary={this._roleListDictionary}
                showEXTERNAL={showEXTERNAL}
                onDownloadCertificate={this.onDownloadBodyCertificate.bind(this)}
              />
            }
            {
              <Dialog open={newUserDialogIsOpen}>
                <DialogSurface className={styles.managementDialog}>
                  <DialogBody>
                    <DialogTitle>
                      {strings.NewUsers}
                    </DialogTitle>
                      {newUsersFromExcel ?
                        <NewUsersFromExcel 
                          isMobile={isMobile}
                          onClosePanel={this.onCloseNewUserPanel.bind(this)}
                          onRealoadComponent={this.onReloadItems.bind(this)}
                          organFilterOptions={this._organListWithBodyId}
                          bkService={this.props.bkService}
                          context={this.props.context}
                          invitationMsg={this._invitationMessage}
                        />
                      :
                        <NewUserForm 
                          organFilterOptions={organFilterOptions} 
                          ccaaFilterOptions={ccaaFilterOptions} 
                          roleFilterOptions={roleFilterOptions}
                          onCreateUser={this.onSaveNewUserPanel.bind(this)}
                          onClosePanel={this.onCloseNewUserPanel.bind(this)}
                          hasErrorNewUser={errorCreatingUser}
                          bkService={this.props.bkService}
                          invitationMsg={this._invitationMessage}
                        />
                      }
                  </DialogBody>
                </DialogSurface>
              </Dialog >
            }
            {
              <Dialog open={dialogIsOpen}>
                <DialogSurface className={styles.managementDialog}>
                  <DialogBody>
                    <DialogTitle>
                      {searchType === SearchType.Organ ? strings.ManageOrgan : strings.ManageUser}
                    </DialogTitle>
                    <DepartmentDetails
                      managedItem={managedItem}
                      onManageItem={this.onManageItem.bind(this)}
                      organTypeFilterOptions={organTypeFilterOptions}
                      ministryFilterOptions={ministryFilterOptions}
                      organFilterOptions={organFilterOptions}
                      ccaaFilterOptions={ccaaFilterOptions}
                      roleFilterOptions={roleFilterOptions}
                      legistatureOptions={legistatureOptions}
                      secretariaOptions={secretariaOptions}
                      BusinessAreaOptions={BusinessAreaOptions}
                      isMobile={isMobile}
                      userDictionary={this._usersListDictionary}
                      roleDictionary={this._roleListDictionary}
                      organDictionary={this._organListDictionary}
                      organTypeDictionary={this._organTypeListDictionary}
                      ministryDictionary={this._ministryDictionary}
                      errorUpdating={errorUpdating}
                      onCheckPermissions={this.onCheckPermissions.bind(this)}
                      licensesInformation={availableLicenses}
                      onUpdateLicense={this.onUpdateLicense.bind(this)}
                      showMode={componentTypeMode.All}
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