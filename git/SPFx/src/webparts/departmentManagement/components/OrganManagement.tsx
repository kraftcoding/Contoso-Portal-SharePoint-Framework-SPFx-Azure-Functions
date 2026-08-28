import * as React from 'react';
import styles from './departmentManagement.module.scss';
import type { IdepartmentManagementprops, IdepartmentManagementState } from './IdepartmentManagement';
import { BodyUserInfo, componentTypeMode, IOrganResult, IPermissionCheckResult, IUserPermissionsCheck, IUserResult, IUsersOrganRelation, OrganTaxonomyIds, RoleTaxonomyIds, SearchType } from '../../administrationApp/models/AdministrationAppModels';
import DepartmentDetails from '../../../components/DepartmentDetails/DepartmentDetails';
import { Button, Dialog, DialogBody, DialogSurface, DialogTitle, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Spinner } from '@fluentui/react-components';
import { getBodyIdFormUrl } from '../../../service/RoleService';
import { Term } from '../../../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy';
import strings from 'departmentManagementWebPartStrings';
import NewUsersFromExcel from '../../../components/NewUsersFromExcel/NewUsersFromExcel';
import NewUserForm from '../../../components/NewUserForm/NewUserForm';
import { PeopleAddRegular, PeopleCommunityAddRegular, PersonAddRegular } from '@fluentui/react-icons';
import { serializeError } from '../../../utils/errorUtils';

export default class departmentManagement extends React.Component<IdepartmentManagementprops, IdepartmentManagementState> {

  private _usersListDictionary: Record<string, IUserResult>;
  private _organListDictionary: Record<string, string>;
  private _organTypeListDictionary: Record<string, string>;
  private _ministryDictionary: Record<string, string>;
  //private _communityListDictionary: Record<string, string>;
  private _roleListDictionary: Record<string, string>;
  private _invitationMessage: string;
  private _organListWithBodyId: Term[];

  constructor(props: IdepartmentManagementprops) {
    super(props);
    this.state = {
      isLoading: true,
      organTypeOptions: [],
      ministryOptions: [],
      Departmentptions: [],
      roleOptions: [],
      secretariaOptions: [],
      legistatureOptions: [],
      ccaaFilterOptions: [],
      managedItem: {
        itemId: "",
        itemType: SearchType.Organ,
        organDialogContent: undefined,
        userDialogContent: undefined
      },
      organBackList: [],
      userBackList: [],
      BusinessAreaOptions: [],
      showLoadDataError: false,
      loadDataError: null,
      isLoadingUpdates: true,
      errorUpdating: false,
      newUserDialogIsOpen:false,
      newUsersFromExcel:false,
      errorCreatingUser:false
    };
  }

  async componentDidMount(): promise<void> {
    await this.loadData();
    void this.getTaxonomies();
    this.props.bkService.getUserInvitationMessage("/sites/Contoso/NotificationUserConfig/Messages/InvitacionUsuarios.txt").then(value => {
      this._invitationMessage = value
    }).catch(ex => console.log(ex))
 }

  private loadData(): void {
    const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
    const request: promise<any>[] = [];
    request.push(
      this.props.bkService.getOrganInfo(bodyId),
      this.props.bkService.getUsersInfo(bodyId)
    )

    promise.all(request).then(results => {
      const organs: IOrganResult[] = [];
      let users: Record<string, IUserResult> = {};
      organs.push(...results[0]);
      users = results[1];
      const orderedOrgans = organs.sort((a, b) => this._organListDictionary[a.Department].localeCompare(this._organListDictionary[b.Department]));
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
        
      this.setState({ 
        Departmentptions:organsFilterFromList,
        organBackList: orderedOrgans, 
        userBackList: orderedUsers, 
        isLoading: false,
        managedItem: {
          itemId: organs && organs.length > 0 ? organs[0]?.BodyId : "",
          itemType: SearchType.Organ,
          organDialogContent: organs && organs.length > 0 ? organs[0] : undefined,
          userDialogContent: undefined,
        } 
      });
    }).catch(ex => {
      this.setState({ showLoadDataError: true, isLoading: false, loadDataError: serializeError(ex) });
      console.error(ex);
    });
  }


    private onCloseNewUserPanel():void{
      this.setState({
        newUserDialogIsOpen: false
      })
    }
  
    private onReloadItems(): void{
      this.setState({ managedItem: undefined, isLoading: true, isLoadingUpdates: false, newUserDialogIsOpen: false });
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
          this.setState({ managedItem: undefined, isLoading: true, isLoadingUpdates: false, newUserDialogIsOpen: false });
          this.loadData();
        }
      ).catch(ex => {
        console.log(ex);
        this.setState({ isLoadingUpdates: false, errorCreatingUser: true });
      })
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

    this._roleListDictionary = usersRoleTermInfo.reduce((acc, item) => {
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
    /*
        this._communityListDictionary = usersCommunityTermInfo.reduce((acc, item) => {
          acc[item.key] = item.text;
          return acc;
        }, {} as Record<string, string>);
    
    */
    this.setState({
      organTypeOptions: organsTypeTermInfo,
      ministryOptions: organsMinistryTermInfo,
      //Departmentptions: organsTermInfo,
      roleOptions: usersRoleTermInfo,
      ccaaFilterOptions: usersCommunityTermInfo,
      secretariaOptions: organsSecretariaTermInfo,
      legistatureOptions: organsPeriodTermInfo,
      BusinessAreaOptions: BusinessAreaTermInfo
    });
  }

  private async onCheckPermissions(item: IUserPermissionsCheck): promise<IPermissionCheckResult> {
      let bodyId = item.BodyId;
      if (item.BodyId) {
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
      console.log(itemToCheck)
      return this.props.bkService.checkUserHasPermissions(itemToCheck);
    }

  private async onUpdateLicense(upn: string, license: string): promise<void>{
    const mypromise = new promise<void>((resolve) => {
      setTimeout(() => {
        resolve();
      }, 3000);
    });
    return mypromise;
  }

  private async onManageItem(itemId: string, itemType: SearchType, action: string, item?: IOrganResult | IUserResult | IUsersOrganRelation[]): promise<void> {
    const { bkService } = this.props;
    if (action === "Open") {
      this.setState({
        // dialogIsOpen: true,
        managedItem: {
          itemId,
          itemType,
          organDialogContent: this.state.managedItem?.organDialogContent,
          userDialogContent: undefined
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
        //dialogIsOpen: false,
        managedItem: {
          itemId,
          itemType,
          organDialogContent: this.state.managedItem?.organDialogContent,
          userDialogContent: undefined
        },
        newUserDialogIsOpen: false,
        errorCreatingUser: false
      });
    }
    else if (action === "Updateproperties") {
      const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
      console.log(item);
      if (item) {
        if (itemType === SearchType.Organ) {
          const updateItem: IOrganResult = item as IOrganResult;
          this.setState({
            isLoadingUpdates: true,
            errorUpdating: false
          });

          bkService.updateOrganInfo(updateItem, bodyId ? bodyId : 'Contoso').then(() => {
            this.setState({ managedItem: undefined, isLoading: true, isLoadingUpdates: false });
            this.loadData();
          }
          ).catch(ex => {
            console.log(ex);
            this.setState({ isLoadingUpdates: false, errorUpdating: true });
          })
          console.log(updateItem);

        } else if (itemType === SearchType.User) {
          this.setState({
            isLoadingUpdates: true,
            errorUpdating: false
          });
          const updateItem: IUserResult = item as IUserResult;
          bkService.updateUserprofile(updateItem, bodyId ? bodyId : 'Contoso').then(() => {
            this.setState({ managedItem: undefined, isLoading: true, isLoadingUpdates: false });
            this.loadData();
          }
          ).catch(ex => {
            console.log(ex);
            this.setState({ isLoadingUpdates: false, errorUpdating: true });
          })
        }
      }
    } else if (action === "UpdateMember") {
      console.log(item);

      if (item) {
        const bodyId: string = getBodyIdFormUrl(this.props.context.pageContext.site.serverRelativeUrl);
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
        bkService.updateUsersInOrgan(organUserRelation, bodyId ? bodyId: 'Contoso').then(() => {
          this.setState({ managedItem: undefined, isLoading: true, isLoadingUpdates: false});
          this.loadData();
        }
        ).catch(ex => {
          console.log(ex);
          this.setState({ isLoadingUpdates: false, errorUpdating: true});
        })
      }
    }
  }

  public render(): React.ReactElement<IdepartmentManagementprops> {
    const {showLoadDataError, loadDataError, ccaaFilterOptions, newUserDialogIsOpen, newUsersFromExcel, errorCreatingUser, isLoading, legistatureOptions, BusinessAreaOptions, secretariaOptions, managedItem, roleOptions, Departmentptions, ministryOptions, organTypeOptions } = this.state;
    const { isMobile, componentTitle } = this.props;
    return (
      <div className={`${styles.departmentManagement}`}>
        <div className={styles.wpTitle}>
          {
          (componentTitle && componentTitle !== "") ?
            <div>{this.props.componentTitle}</div>
            :
            <div></div>
          }
          {!isLoading && !showLoadDataError &&
            <div className={styles.newElementsButtons}>
              {isMobile ? 
                <>
                <Button appearance={"primary"} icon={<PersonAddRegular />} onClick={() => this.onManageItem("", SearchType.Organ, "OpenNew")}>{strings.NewUser}</Button>
                <Button appearance={"primary"} icon={<PeopleCommunityAddRegular />} onClick={() => this.onManageItem("", SearchType.Organ, "OpenNewMultiple")}>{strings.NewMultipleUsers}</Button>
                </>
              :
                <Menu>
                  <MenuTrigger>
                      <Button appearance={"primary"} icon={<PeopleAddRegular />}>{strings.NewUsers}</Button>
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem icon={<PersonAddRegular />} onClick={() => this.onManageItem("", SearchType.Organ, "OpenNew")}>{strings.NewUser}</MenuItem>
                      <MenuItem icon={<PeopleCommunityAddRegular />} onClick={() => this.onManageItem("", SearchType.Organ, "OpenNewMultiple")}>{strings.NewMultipleUsers}</MenuItem>
                    </MenuList>
                  </MenuPopover>
                </Menu>
              }
            </div>
          }
        </div>
        {
          <Dialog open={newUserDialogIsOpen} >
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
                      testSelectedOrgan={this._organListWithBodyId && this._organListWithBodyId.length>0 ? this._organListWithBodyId[0]: undefined}
                      invitationMsg={this._invitationMessage}
                    />
                  :
                    <NewUserForm 
                      organFilterOptions={Departmentptions} 
                      ccaaFilterOptions={ccaaFilterOptions} 
                      roleFilterOptions={roleOptions}
                      onCreateUser={this.onSaveNewUserPanel.bind(this)}
                      onClosePanel={this.onCloseNewUserPanel.bind(this)}
                      hasErrorNewUser={errorCreatingUser}
                      testSelectedOrgan={Departmentptions && Departmentptions.length>0 ? Departmentptions[0]: undefined}
                      bkService={this.props.bkService}
                      invitationMsg={this._invitationMessage}
                    />
                  }
              </DialogBody>
            </DialogSurface>
          </Dialog >
        }
        {!isLoading ?
        <>
          {showLoadDataError ?
            <div>
                <div className={styles.errorLabel}>
                  {strings.LoadingError}
                  {loadDataError && <test style={{ whiteSpace: 'test-wrap', margin: '8px 0 0', fontSize: '12px', fontFamily: 'monospace', wordBreak: 'break-all' }}>{loadDataError}</test>}
                </div>
            </div>
          :
       
          <DepartmentDetails
            managedItem={managedItem}
            onManageItem={this.onManageItem.bind(this)}
            organTypeFilterOptions={organTypeOptions}
            ministryFilterOptions={ministryOptions}
            organFilterOptions={Departmentptions}
            ccaaFilterOptions={[]}
            roleFilterOptions={roleOptions}
            legistatureOptions={legistatureOptions}
            secretariaOptions={secretariaOptions}
            isMobile={isMobile}
            userDictionary={this._usersListDictionary}
            roleDictionary={this._roleListDictionary}
            organDictionary={this._organListDictionary}
            organTypeDictionary={this._organTypeListDictionary}
            ministryDictionary={this._ministryDictionary}
            BusinessAreaOptions={BusinessAreaOptions}
            onCheckPermissions={this.onCheckPermissions.bind(this)}
            onUpdateLicense={this.onUpdateLicense.bind(this)}
            showMode={componentTypeMode.Users}
            showConfirmOnPanel={true}
          />
        }
        </>
          :
          <Spinner size='large' labelPosition="below" label={strings.LoadingLabel} />
        }
      </div>
    );
  }
}
