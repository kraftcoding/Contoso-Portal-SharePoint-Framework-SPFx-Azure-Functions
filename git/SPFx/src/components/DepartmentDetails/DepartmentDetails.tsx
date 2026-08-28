import * as React from 'react';
import { IDepartmentDetailsState } from './IDepartmentDetails';
import { IDepartmentDetailsprops } from './IDepartmentDetails';
import styles from './DepartmentDetails.module.scss';
import { componentTypeMode, IErrorsForm, IMemberOrgan, IObjectDifferences, IOrganResult, IUserOrganList, IUserPermissionsCheck, IUserResult, IUsersOrganRelation, LicenseNames, LicensesTypes, OrganTaxonomyIds, RoleTaxonomyIds, SearchType } from '../../webparts/administrationApp/models/AdministrationAppModels';
import {
  Button,
  DialogActions,
  DialogTrigger,
  Divider,
  Dropdown,
  Input,
  Label,
  Switch,
  Tab,
  TabList,
  Textarea,
  Option,
  Combobox,
  DialogContent,
  Spinner,
  Tooltip,
  //Dialog,
  //DialogBody,
  //DialogTitle,
  Popover,
  PopoverSurface,
  PopoverTrigger,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle
} from '@fluentui/react-components';
import {
  ContactCardRibbon20Regular,
  DeleteRegular,
  Dismiss20Regular,
  ErrorCircle20Filled,
  Info12Filled,
  //Folder20Regular,
  ListBar20Regular,
  PeopleTeam20Regular,
  Person20Regular,
  PersonKey20Regular,
  Settings20Regular,
  Warning20Regular
} from '@fluentui/react-icons';
import _ from 'lodash';
import format from 'date-fns/format';
import { PeoplePicker, PersonType, UserType, ViewType } from '@microsoft/mgt-react';
import { File, /* FileList,*/ MgtTemplateprops } from '@microsoft/mgt-react/dist/es6/spfx';
import { DriveItem } from '@microsoft/microsoft-graph-types';
import strings from 'DepartmentDetailsStrings';

export const FileTemplate = (props: MgtTemplateprops & IDepartmentDetailsprops & { onDeleteClick: any }): any => {
  const file: DriveItem | undefined = props.dataContext ? props.dataContext.file : undefined;

  //este trozo pendiente hasta que podamos hacer un column-reverse del mgt-file-list
  /*const estilo = document.createElement("style");
  estilo.innerHTML = ".file-list-wrapper{flex-direction: column-reverse;}"
  const elemento = document.querySelector(`.${styles.fileList}`);
  elemento?.shadowRoot?.appendChild(estilo);*/

  //el boton de borrado de ficheros solo estara disponible para estos estados
  //const estadoConBorrado = props.event.StatusId === EventStatus.InConstruction || props.event.StatusId === EventStatus.testBooking || props.event.StatusId === EventStatus.Published;
  const estadoConBorrado = true;
  return (
    file?.name?.indexOf("_hidden") === -1 ?
      <>
        <div style={{ display: 'flex', flexDirection: 'row' }}>
          <File view={ViewType.twolines} fileDetails={file} onClick={(props) => {
            if (file && !file.folder) {
              window.open(file.webUrl?.toString(), "_blank", "noreferrer")
            }
          }}></File>
          {
            estadoConBorrado && !file.folder &&
            <div style={{ flexGrow: 1, display: 'flex', alignSelf: 'center', justifyContent: 'flex-end' }}>
              <Button title={"Borrar"} appearance='secondary' icon={<DeleteRegular />} onClick={() => { props.onDeleteClick(file) }} />
            </div>
          }
        </div>
        <Divider appearance='subtle'></Divider>
      </>
      :
      <></> //no pintamos el item
  );
}

class DepartmentDetails extends React.Component<IDepartmentDetailsprops, IDepartmentDetailsState> {
  private indexKeyNewListItem: number = 0;

  constructor(props: IDepartmentDetailsprops) {
    super(props);
    this.state = {
      selectedTab: this.props.showMode === componentTypeMode.Users ? "Listado"  : "prodpiedades",
      editSelectedItem: this.props.managedItem,
      newValues: {
        Type: this.props.managedItem.itemType,
        Items: [],
        hasError: false
      },
      showConfirmation: false,
      showLoadingConfirmation: false,
      formErrors: {},
      showCheckPermissionsModal: false,
      checkPermissionsInfo: undefined
    };
  }
  componentDidUpdate(testvprops: Readonly<IDepartmentDetailsprops>): void {
    if (!_.isEqual(testvprops.managedItem, this.props.managedItem)) {
      this.setState({
        editSelectedItem: this.props.managedItem,
        newValues: {
          Type: this.props.managedItem ? this.props.managedItem.itemType : undefined,
          Items: [],
          hasError: false
        },
        selectedTab: "prodpiedades",
        showConfirmation: false,
        showLoadingConfirmation: false,
        formErrors: {},
        showCheckPermissionsModal: false,
        checkPermissionsInfo: undefined,
        currentLicense: undefined,
        //showErrorUpdate: false
      });
      this.indexKeyNewListItem = 0;
    }
  }

  private onClickClose(): void {
    const { managedItem, onManageItem, showConfirmOnPanel } = this.props;
    onManageItem(managedItem.itemId, managedItem.itemType, "Close")
    if(showConfirmOnPanel)
    this.setState({
      editSelectedItem: this.props.managedItem,
      newValues: {
        Type: this.props.managedItem ? this.props.managedItem.itemType : undefined,
        Items: [],
        hasError: false
      },
      showConfirmation: false 
    })
  }

  private onClickSave(): void {
    const { selectedTab, editSelectedItem } = this.state;
    if (selectedTab === "prodpiedades") {
      const hasErrors = this.onCheckValuesForm();
      console.log(hasErrors);
      if (!hasErrors && !editSelectedItem?.isNewItem) {
        this.setState({ showConfirmation: true });
      }
    } else if (selectedTab === 'Listado') {
      /*const modifiedListItems: IUsersOrganRelation[] = this.onGetListDifferences();
      console.log(modifiedListItems);*/
      const hasDuplicates = this.onCheckDuplicates();
      if (hasDuplicates) {
        this.setState({ newValues: { ...this.state.newValues, hasError: true } });
      } else {
        this.setState({ showConfirmation: true });
      }
    }else if (selectedTab === 'Licencias'){
      this.setState({ showConfirmation: true });
    }
  }

  private onCheckDuplicates(): boolean {
    const { editSelectedItem, newValues } = this.state;
    let hasDuplicatedItems = false;
    if (editSelectedItem.itemType === SearchType.Organ) {
      const updatedUsers: IMemberOrgan[] = editSelectedItem?.organDialogContent?.Users;
      newValues.Items.forEach(newItem => {
        if (!hasDuplicatedItems) {
          if (updatedUsers && updatedUsers.some(us => us.UserPrincipalName === newItem.UserId)) {
            hasDuplicatedItems = true;
          } else if (newValues && newValues.Items.some(us => (us.UserId === newItem.UserId && us.IndexKey !== newItem.IndexKey))) {
            hasDuplicatedItems = true;
          }
        }
      });

    } else if (editSelectedItem.itemType === SearchType.User) {
      const updatedOrgans: IUserOrganList[] = editSelectedItem?.userDialogContent?.Organ;
      newValues.Items.forEach(newItem => {
        if (!hasDuplicatedItems) {
          if (updatedOrgans && updatedOrgans.some(or => or.OrganId === newItem.OrganId)) {
            hasDuplicatedItems = true;
          } else if (newValues && newValues.Items.some(or => (or.OrganId === newItem.OrganId && or.IndexKey !== newItem.IndexKey))) {
            hasDuplicatedItems = true;
          }
        }
      });
    }
    return hasDuplicatedItems;
  }

  private onGetListDifferences(): IUsersOrganRelation[] {
    const { editSelectedItem, newValues } = this.state;
    const { managedItem } = this.props;
    const itemsToModify: IUsersOrganRelation[] = [];
    if (managedItem.itemType === SearchType.Organ) {
      const updatedUsers: IMemberOrgan[] = editSelectedItem?.organDialogContent?.Users;
      const originalUsers: IMemberOrgan[] = managedItem?.organDialogContent.Users;
      if (originalUsers && updatedUsers) {

        originalUsers.forEach(item => {
          const isUserModified = updatedUsers.find(us => (us.UserPrincipalName === item.UserPrincipalName && us.RoleId !== item.RoleId));
          if (isUserModified) {
            itemsToModify.push({
              UserUpn: isUserModified.UserPrincipalName?.toLocaleLowerCase(),
              BodyId: managedItem?.organDialogContent?.BodyId,
              Role: isUserModified.RoleId
            });
          }
          const isUserRemoved = !updatedUsers.some(us => us.UserPrincipalName === item.UserPrincipalName);
          if (isUserRemoved) {
            itemsToModify.push({
              UserUpn: item.UserPrincipalName?.toLocaleLowerCase(),
              BodyId: managedItem?.organDialogContent?.BodyId,
              Role: undefined
            });
          }
        });
      }
      if (newValues) {
        newValues.Items.forEach(newItem => {
          if (newItem.UserId && newItem.Role) {
            itemsToModify.push({
              UserUpn: newItem.UserId?.toLocaleLowerCase(),
              BodyId: managedItem?.organDialogContent?.BodyId,
              Role: newItem.Role
            })
          }
        })
      }
    }
    else if (managedItem.itemType === SearchType.User) {
      const updatedOrgans: IUserOrganList[] = editSelectedItem?.userDialogContent?.Organ;
      const originalOrgans: IUserOrganList[] = managedItem?.userDialogContent?.Organ;
      if (updatedOrgans && originalOrgans) {
        originalOrgans.forEach(item => {
          const organModified = updatedOrgans.find(or => (or.OrganId === item.OrganId && or.RoleId !== item.RoleId));
          if (organModified && organModified.OrganId) {
            itemsToModify.push({
              UserUpn: managedItem?.userDialogContent?.UserPrincipalName?.toLocaleLowerCase(),
              BodyId: organModified.OrganId,
              Role: organModified.RoleId
            });
          }
          const isUserRemoved = !updatedOrgans.some(or => or.OrganId === item.OrganId);
          if (isUserRemoved && item.OrganId) {
            itemsToModify.push({
              UserUpn: managedItem?.userDialogContent?.UserPrincipalName?.toLocaleLowerCase(),
              BodyId: item.OrganId,
              Role: undefined
            });
          }
        });
      }
      if (newValues) {
        newValues.Items.forEach(newItem => {
          if (newItem.Role && newItem.OrganId) {
            console.log(newItem.Role)
            itemsToModify.push({
              UserUpn: managedItem?.userDialogContent?.UserPrincipalName?.toLocaleLowerCase(),
              BodyId: newItem.OrganId,
              Role: newItem.Role
            })
          }
        })
      }
    }
    return itemsToModify;
  }

  private onClickSaveConfirmation(): void {
    const { managedItem, onManageItem, onUpdateLicense } = this.props;
    const { editSelectedItem, selectedTab, currentLicense } = this.state;
    this.setState({ showLoadingConfirmation: true })
    if (selectedTab === "prodpiedades") {
      if (managedItem.itemType === SearchType.User) {
        onManageItem(managedItem.itemId, managedItem.itemType, "Updateproperties", editSelectedItem?.userDialogContent);
      } else if (managedItem.itemType === SearchType.Organ) {
        let organUpdatedItem: IOrganResult = editSelectedItem?.organDialogContent;
        if (organUpdatedItem.NombreDepartment === managedItem?.organDialogContent?.NombreDepartment) {
          organUpdatedItem.NombreDepartment = ""
        }
        // const newItem = { ...editSelectedItem, organDialogContent: organUpdatedItem }
        onManageItem(managedItem.itemId, managedItem.itemType, "Updateproperties", organUpdatedItem);
      }
      //this.setState({ showConfirmation: false });
    } else if (selectedTab === "Listado") {
      const differences = this.onGetListDifferences();
      onManageItem(managedItem.itemId, managedItem.itemType, "UpdateMember", differences);
    } else if (selectedTab === "Licencias") {
      if(editSelectedItem?.userDialogContent){
        const user = editSelectedItem.userDialogContent as IUserResult;
        onUpdateLicense( user.UserPrincipalName ,(!currentLicense || currentLicense === "nolicense")? "": currentLicense)
      }
    }
  }

  private onClickCloseConfirmation(): void {
    if (this.props.errorUpdating) {
      const { managedItem, onManageItem } = this.props;
      onManageItem(managedItem.itemId, managedItem.itemType, "Close")
    }
    this.setState({ showConfirmation: false });
  }

  private onClickCheckPermissions(upn: string, displayName: string, bodyId: string, role: string): void {
    this.setState({ showCheckPermissionsModal: true, checkPermissionsInfo: { isLoading: true, UserPrincipalName: upn, BodyId: bodyId } });
    const userPermissionsInfo: IUserPermissionsCheck = {
      UserPrincipalName: upn,
      DisplayName: displayName,
      BodyId: bodyId,
      Role: role,
    }
    this.props.onCheckPermissions(userPermissionsInfo)
      .then(result => {
        this.setState({
          checkPermissionsInfo: {
            ...userPermissionsInfo,
            HasPermissions:    result.HasPermissions,
            AssignmentStatus:  result.AssignmentStatus,
            NumIntentos:       result.NumIntentos,
            IsErrorPermanente: result.IsErrorPermanente,
            ErrorMensaje:      result.ErrorMensaje,
            isLoading: false
          }
        });
      }).catch(ex => {
        this.setState({ checkPermissionsInfo: undefined, showCheckPermissionsModal: false });
        console.log(ex);
      });
  }
  /*
  private getObjectDiff(obj1: any, obj2: any) {
    const diff = Object.keys(obj1).reduce((result, key) => {

      if (!Object.prodtotype.hasOwnproperty.call(obj2, key)) {
        result.push(key);
      } else if (_.isEqual(obj1[key], obj2[key])) {
        const resultKeyIndex = result.indexOf(key);
        result.splice(resultKeyIndex, 1);
      }
      return result;
    }, Object.keys(obj2))
    return diff;
  }
*/
  private onGetChanges(): IObjectDifferences[] {
    const { selectedTab } = this.state;
    const currentItem = this.props.managedItem;
    const cambios: IObjectDifferences[] = [];

    if (selectedTab === "prodpiedades") {
      if (currentItem.itemType === SearchType.Organ) {
        const currentItemContent: IOrganResult = currentItem?.organDialogContent;
        const updatedItem = this.state.editSelectedItem?.organDialogContent as IOrganResult;
        if (currentItemContent.NombreDepartment !== updatedItem.NombreDepartment) {
          cambios.push({ property: strings.DepartmentName, oldValue: this.props.organDictionary[currentItemContent.Department], newValue: updatedItem.NombreDepartment ? updatedItem.NombreDepartment : "" })
        }
        if (currentItemContent.TipoDepartment !== updatedItem.TipoDepartment) {
          cambios.push({ property: strings.OrganType, oldValue: this.props.organTypeDictionary[currentItemContent.TipoDepartment], newValue: this.props.organTypeDictionary[updatedItem.TipoDepartment] })
        }
        if (currentItemContent.BusinessArea !== updatedItem.BusinessArea) {
          const currentObject = currentItemContent.BusinessArea && this.props.BusinessAreaOptions.find(opt => opt.key === currentItemContent.BusinessArea) ? this.props.BusinessAreaOptions.find(opt => opt.key === currentItemContent.BusinessArea)?.text : "";
          const newObject = updatedItem.BusinessArea && this.props.BusinessAreaOptions.find(opt => opt.key === updatedItem.BusinessArea) ? this.props.BusinessAreaOptions.find(opt => opt.key === updatedItem.BusinessArea)?.text : "";

          cambios.push({ property: strings.BusinessArea, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
        }
        if (currentItemContent.Division !== updatedItem.Division) {
          cambios.push({ property: strings.Ministry, oldValue: this.props.ministryDictionary[currentItemContent.Division], newValue: this.props.ministryDictionary[updatedItem.Division] })
        }
        if (currentItemContent.Abreviatura !== updatedItem.Abreviatura) {
          cambios.push({ property: strings.Abbreviation, oldValue: currentItemContent.Abreviatura, newValue: updatedItem.Abreviatura })
        }
        if (currentItemContent.BodyId !== updatedItem.BodyId) {
          cambios.push({ property: strings.Abbreviation, oldValue: currentItemContent.BodyId, newValue: updatedItem.BodyId })
        }
        if (currentItemContent.Secretaria !== updatedItem.Secretaria) {
          const currentObject = currentItemContent.Secretaria && this.props.secretariaOptions.find(opt => opt.key === currentItemContent.Secretaria) ? this.props.secretariaOptions.find(opt => opt.key === currentItemContent.Secretaria)?.text : "";
          const newObject = updatedItem.Secretaria && this.props.secretariaOptions.find(opt => opt.key === updatedItem.Secretaria) ? this.props.secretariaOptions.find(opt => opt.key === updatedItem.Secretaria)?.text : "";

          cambios.push({ property: strings.Secretaria, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
        }
        if (currentItemContent.Intersectorial !== updatedItem.Intersectorial) {
          cambios.push({ property: strings.Intersectorial, oldValue: currentItemContent.Intersectorial ? strings.IntersectorialYes : strings.IntersectorialNo, newValue: updatedItem.Intersectorial ? strings.IntersectorialYes : strings.IntersectorialNo })
        }
        if (currentItemContent.Period !== updatedItem.Period) {
          const currentObject = currentItemContent.Period && this.props.legistatureOptions.find(opt => opt.key === currentItemContent.Period) ? this.props.legistatureOptions.find(opt => opt.key === currentItemContent.Period)?.text : "";
          const newObject = updatedItem.Period && this.props.legistatureOptions.find(opt => opt.key === updatedItem.Period) ? this.props.legistatureOptions.find(opt => opt.key === updatedItem.Period)?.text : "";

          cambios.push({ property: strings.Legislature, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
        }
        if (currentItemContent.FechaConstitucion !== updatedItem.FechaConstitucion) {
          cambios.push({ property: strings.CreationDate, oldValue: currentItemContent?.FechaConstitucion ? format(new Date(currentItemContent.FechaConstitucion), 'yyyy-MM-dd') : "", newValue: updatedItem?.FechaConstitucion ? format(new Date(updatedItem.FechaConstitucion), 'yyyy-MM-dd') : "" })
        }
        if (currentItemContent.FechaExtincion !== updatedItem.FechaExtincion) {
          cambios.push({ property: strings.ExpirationDate, oldValue: currentItemContent?.FechaExtincion ? format(new Date(currentItemContent.FechaExtincion), 'yyyy-MM-dd') : "", newValue: updatedItem?.FechaExtincion ? format(new Date(updatedItem.FechaExtincion), 'yyyy-MM-dd') : "" })
        }
        if (currentItemContent.DiasAprodbacionMinutes !== updatedItem.DiasAprodbacionMinutes) {
          cambios.push({ property: strings.DaysToApprodve, oldValue: currentItemContent.DiasAprodbacionMinutes ? currentItemContent.DiasAprodbacionMinutes.toString() : "", newValue: updatedItem.DiasAprodbacionMinutes ? updatedItem.DiasAprodbacionMinutes.toString() : "" })
        }
        if (currentItemContent.SIA !== updatedItem.SIA) {
          cambios.push({ property: strings.SIA, oldValue: currentItemContent.SIA, newValue: updatedItem.SIA })
        }
        if (currentItemContent.dir !== updatedItem.dir) {
          cambios.push({ property: strings.DIR3, oldValue: currentItemContent.dir, newValue: updatedItem.dir })
        }
        if (currentItemContent.Description !== updatedItem.Description) {
          cambios.push({ property: strings.Description, oldValue: currentItemContent.Description, newValue: updatedItem.Description })
        }
        if (currentItemContent.Observaciones !== updatedItem.Observaciones) {
          cambios.push({ property: strings.Observations, oldValue: currentItemContent.Observaciones, newValue: updatedItem.Observaciones })
        }
      } else {
        const currentItemContent: IUserResult = currentItem?.userDialogContent;
        const updatedItem = this.state.editSelectedItem?.userDialogContent as IUserResult;
        //const diff = this.getObjectDiff(currentItemContent, updatedItem);
        if (currentItemContent.DisplayName !== updatedItem.DisplayName) {
          cambios.push({ property: strings.DisplayName, oldValue: currentItemContent.DisplayName, newValue: updatedItem.DisplayName })
        }
        if (currentItemContent.FirstName !== updatedItem.FirstName) {
          cambios.push({ property: strings.Name, oldValue: currentItemContent.FirstName, newValue: updatedItem.FirstName })
        }
        if (currentItemContent.LastName !== updatedItem.LastName) {
          cambios.push({ property: strings.SurName, oldValue: currentItemContent.LastName, newValue: updatedItem.LastName })
        }
        if (currentItemContent.JobTitle !== updatedItem.JobTitle) {
          cambios.push({ property: strings.Job, oldValue: currentItemContent.JobTitle, newValue: updatedItem.JobTitle })
        }
        if (currentItemContent.Office !== updatedItem.Office) {
          cambios.push({ property: strings.AutonomousCommunity, oldValue: currentItemContent.Office, newValue: updatedItem.Office })
        }
        if (currentItemContent.CompanyName !== updatedItem.CompanyName) {
          cambios.push({ property: strings.CompanyNameLabel, oldValue: currentItemContent.CompanyName, newValue: updatedItem.CompanyName })
        }
        if (currentItemContent.Department !== updatedItem.Department) {
          cambios.push({ property: strings.DepartmentLabel, oldValue: currentItemContent.Department, newValue: updatedItem.Department })
        }
        if (currentItemContent.EmployeeType !== updatedItem.EmployeeType) {
          cambios.push({ property: strings.EmployeeTypeLabel, oldValue: currentItemContent.EmployeeType, newValue: updatedItem.EmployeeType })
        }
        if (currentItemContent.Email !== updatedItem.Email) {
          cambios.push({ property: strings.SecondaryMail, oldValue: currentItemContent.Email, newValue: updatedItem.Email })
        }
        if (currentItemContent.CellPhone !== updatedItem.CellPhone) {
          cambios.push({ property: strings.PhoneNumber, oldValue: currentItemContent.CellPhone, newValue: updatedItem.CellPhone })
        }
        if (currentItemContent.BusinessPhone !== updatedItem.BusinessPhone) {
          cambios.push({ property: strings.BusinessPhoneLabel, oldValue: currentItemContent.BusinessPhone, newValue: updatedItem.BusinessPhone })
        }
      }
    } else if (selectedTab === "Listado") {
      const { userDictionary, organDictionary, roleDictionary } = this.props;
      const modifiedListItems: IUsersOrganRelation[] = this.onGetListDifferences();
      if (currentItem.itemType === SearchType.Organ) {
        modifiedListItems.forEach(item =>
          cambios.push({ property: userDictionary[item.UserUpn]?.DisplayName, oldValue: "", newValue: item.Role ? roleDictionary[item.Role] : "" })
        )
      } else if (currentItem.itemType === SearchType.User) {
        modifiedListItems.forEach(item =>
          cambios.push({ property: organDictionary[item?.BodyId], oldValue: "", newValue: item.Role ? roleDictionary[item.Role] : "" })
        )
      }
    }
    return cambios;
  }

  private onCheckValuesForm(): boolean {
    const { editSelectedItem } = this.state;
    let hasErrors = false;
    let errors: IErrorsForm = {};

    if (editSelectedItem && editSelectedItem?.itemType === SearchType.Organ) {
      const updatedItem = editSelectedItem?.organDialogContent as IOrganResult;
      if (!updatedItem.NombreDepartment) {
        hasErrors = true;
        errors.Department = strings.FieldRequired
      } else {
        const currentItem = this.props.managedItem;
        const currentItemContent: IOrganResult = currentItem?.organDialogContent;
        if (currentItemContent && currentItemContent.NombreDepartment && currentItemContent.NombreDepartment !== updatedItem.NombreDepartment) {
          if (this.props.organFilterOptions.some(or => or?.text?.toLowerCase().trim() === updatedItem.NombreDepartment.toLowerCase().trim())) {
            hasErrors = true;
            errors.Department = strings.DepartmentNameExist
          }
        }
      }
      /*if (!updatedItem.Division || updatedItem.Division === OrganTaxonomyIds.NullTaxonomy) {
        hasErrors = true;
        errors.Division = strings.FieldRequired
      }
        */
      if (!updatedItem.BusinessArea || updatedItem.BusinessArea === OrganTaxonomyIds.NullTaxonomy) {
        hasErrors = true;
        errors.BusinessArea = strings.FieldRequired
      }
      if (!updatedItem.Period || updatedItem.Period === OrganTaxonomyIds.NullTaxonomy) {
        hasErrors = true;
        errors.Period = strings.FieldRequired
      }
      /*if (!updatedItem.SIA) {
        hasErrors = true;
        errors.SIA = strings.FieldRequired
      }
      if (!updatedItem.dir) {
        hasErrors = true;
        errors.dir = strings.FieldRequired
      }*/
    }
    else if (editSelectedItem && editSelectedItem?.itemType === SearchType.User) {
      const updatedItem = editSelectedItem?.userDialogContent as IUserResult;
      if (!updatedItem.DisplayName) {
        hasErrors = true;
        errors.DisplayName = strings.FieldRequired;
      }
      if (!updatedItem.FirstName) {
        hasErrors = true;
        errors.FirstName = strings.FieldRequired;
      }
      if (!updatedItem.LastName) {
        hasErrors = true;
        errors.LastName = strings.FieldRequired;
      }
      if (updatedItem.Email && !updatedItem.Email.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) {
        hasErrors = true;
        errors.Email = strings.MailError;
      }
      if (updatedItem.CellPhone && !updatedItem.CellPhone.match(/^\d{9}$/)) {
        hasErrors = true;
        errors.CellPhone = strings.PhoneError;
      }
    }
    this.setState({ formErrors: errors });
    return hasErrors;
  }

  private onUpdateproperty(type: SearchType, property: string, value: string | number | boolean | undefined): void {
    const { editSelectedItem } = this.state;
    if (type === SearchType.Organ) {
      const organDialogContent: IOrganResult = editSelectedItem?.organDialogContent;

      if (["Department", "TipoDepartment", "BodyId", "Secretaria", "Division", "Materia", "Period", "Abreviatura", "SIA", "dir", "Description", "Observaciones", "NombreDepartment", "BusinessArea"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value } } });
      } else if (["FechaConstitucion", "FechaExtincion"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value } } })
      } else if (["DiasAprodbacionMinutes"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value } } })
      } else if (["Intersectorial"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value } } })
      }

    } else if (type === SearchType.User) {
      const userDialogContent: IUserResult = editSelectedItem?.userDialogContent;
      if (["FirstName", "LastName", "DisplayName", "Email", "OtherMails", "JobTitle", "CellPhone", "Office", "BusinessPhone", "Department", "EmployeeType", "CompanyName"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, userDialogContent: { ...userDialogContent, [property]: value } } })
      } else if (["AutonomousCommunity"].indexOf(property) > -1) {
        this.setState({ editSelectedItem: { ...editSelectedItem, userDialogContent: { ...userDialogContent, [property]: value } } })
      }
    }

  }

  private onUpdateListItem(type: SearchType, action: string, itemId: string, roleValue?: string, ccaaValue?: string): void {
    const { editSelectedItem, showCheckPermissionsModal, checkPermissionsInfo } = this.state;
    if (type === SearchType.Organ) {
      const organDialogContent: IOrganResult = editSelectedItem?.organDialogContent;
      if (action === 'delete') {
        const allUsers = organDialogContent.Users;

        const newUsers = allUsers.filter(user => user.UserPrincipalName !== itemId)
        if (showCheckPermissionsModal && checkPermissionsInfo && checkPermissionsInfo.UserPrincipalName === itemId) {
          this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Users: newUsers } }, showCheckPermissionsModal: false, checkPermissionsInfo: undefined })
        } else {
          this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Users: newUsers } } })
        }
      } else if (action === 'update') {
        const allUsers = organDialogContent.Users;
        const newUsers = roleValue ? allUsers.map(user => user.UserPrincipalName === itemId ? { ...user, RoleId: roleValue } : user) : allUsers.map(user => user.UserPrincipalName === itemId ? { ...user, Ccaa: ccaaValue } : user);
        this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Users: newUsers } } })
      }
    } else if (type === SearchType.User) {
      const userDialogContent: IUserResult = editSelectedItem?.userDialogContent;
      if (action === 'delete') {
        const allOrgans = userDialogContent.Organ;
        const newOrgans = allOrgans?.filter(org => org.OrganId !== itemId)
        this.setState({ editSelectedItem: { ...editSelectedItem, userDialogContent: { ...userDialogContent, Organ: newOrgans } } })
      } else if (action === 'update') {
        const allUsers = userDialogContent.Organ;
        const newUsers = allUsers?.map(org => org.OrganId === itemId ? { ...org, RoleId: roleValue } : org)
        this.setState({ editSelectedItem: { ...editSelectedItem, userDialogContent: { ...userDialogContent, Organ: newUsers } } })
      }
    }
  }

  private onUpdateNewListItems(type: SearchType, action: string, itemId?: string, rol?: string, keyIndex?: number, ccaa?: string): void {
    const { newValues } = this.state;
    const newArray = [...newValues.Items];
    if (type === SearchType.Organ) {
      if (action === 'add') {
        if (newArray.length === 0 || (newArray.length > 0 && newArray[newArray.length - 1].UserId !== undefined && newArray[newArray.length - 1].Role !== undefined)) {
          newArray.push({ UserId: undefined, Role: undefined, IndexKey: this.indexKeyNewListItem + 1 });
          this.setState({ newValues: { ...newValues, Items: newArray } });
          this.indexKeyNewListItem = this.indexKeyNewListItem + 1;
        }
      } else if (action === 'update') {
        if (newArray.length > 0) {
          const updatedArray = newArray.map((x) => x.IndexKey === keyIndex ? rol ? ccaa ? { ...x, UserId: itemId, Role: rol, Ccaa: ccaa } : { ...x, UserId: itemId, Role: rol } : ccaa ? { ...x, UserId: itemId, Ccaa: ccaa } : { ...x, UserId: itemId } : x);
          this.setState({ newValues: { ...newValues, Items: updatedArray } });
        }
      } else if (action === 'delete') {
        if (newArray.length > 0) {
          const updatedArray = newArray.filter((x) => x.IndexKey !== keyIndex);
          this.setState({ newValues: { ...newValues, Items: updatedArray } });
        }
      }
    } else if (type === SearchType.User) {
      if (action === 'add') {
        if (newArray.length === 0 || (newArray.length > 0 && newArray[newArray.length - 1].OrganId !== undefined && newArray[newArray.length - 1].Role !== undefined)) {
          newArray.push({ OrganId: undefined, Role: undefined, IndexKey: this.indexKeyNewListItem + 1 });
          this.setState({ newValues: { ...newValues, Items: newArray } });
          this.indexKeyNewListItem = this.indexKeyNewListItem + 1;
        }
      } else if (action === 'update') {
        if (newArray.length > 0) {
          const updatedArray = newArray.map((x) => x.IndexKey === keyIndex ? { ...x, OrganId: itemId, Role: rol } : x);
          this.setState({ newValues: { ...newValues, Items: updatedArray } });
        }
      } else if (action === 'delete') {
        if (newArray.length > 0) {
          const updatedArray = newArray.filter((x) => x.IndexKey !== keyIndex);
          this.setState({ newValues: { ...newValues, Items: updatedArray } });
        }
      }
    }
  }

  public itemClick(e: CustomEvent<DriveItem>) {
    if (e.detail && e.detail.folder) {
      //let id = e.detail.id || "";
      //let name = e.detail.name || "";
      // let parentReference = e.detail.parentReference || undefined;
      console.log(e.detail);
      console.log(e.detail.folder);

        // render new file list
      /*  const { breadCrumbItems } = this.state;

        breadCrumbItems?.push({ title: name, id, parent: parentReference });

        this.setState({ breadCrumbItems, currentDriveItemId: { title: name, id, parent: parentReference } });
    */}
  }

  public onFileSelect(file: DriveItem | undefined): void {
    console.log(file);
  }

  /*
  private onCheckRoleLicense(roleId: string, upn: string): boolean {
    const { userDictionary } = this.props;
    let needLicense = false;
    if ((roleId === RoleTaxonomyIds.Convocante) || (roleId === RoleTaxonomyIds.GestorConvocante)) {
      if (userDictionary && userDictionary[upn.toLowerCase()]) {
        const userLicenses = userDictionary[upn.toLowerCase()].AssignedLicenses;
        if (!userLicenses.includes(LicensesTypes.E3) && !userLicenses.includes(LicensesTypes.E5) && !userLicenses.includes(LicensesTypes.E5Developer)) {
          needLicense = true
        }
      }
    }
    return needLicense;
  }
    */
  public render(): JSX.Element {
    const {showConfirmOnPanel, showMode, licensesInformation, organDictionary, errorUpdating, roleDictionary, userDictionary, isMobile, organFilterOptions, BusinessAreaOptions, ccaaFilterOptions, roleFilterOptions, ministryFilterOptions, organTypeFilterOptions, secretariaOptions, legistatureOptions } = this.props;
    const { currentLicense, selectedTab, editSelectedItem, newValues, showConfirmation, showLoadingConfirmation, formErrors, checkPermissionsInfo, showCheckPermissionsModal } = this.state;

    const organDialogContent: IOrganResult | undefined = (editSelectedItem?.organDialogContent || (editSelectedItem?.isNewItem && editSelectedItem?.itemType === SearchType.Organ)) ?
      editSelectedItem?.organDialogContent
      :
      undefined;
    const userDialogContent: IUserResult | undefined = (editSelectedItem?.userDialogContent || (editSelectedItem?.isNewItem && editSelectedItem?.itemType === SearchType.User)) ?
      editSelectedItem?.userDialogContent
      :
      undefined;
    let listingTabName: string;
    switch (editSelectedItem && editSelectedItem.itemType) {
      case SearchType.Organ:
        listingTabName = strings.Users;
        break;
      case SearchType.User:
        listingTabName = strings.Organs;
        break;
      default:
        listingTabName = "";
        break;
    }
    if (showConfirmation && !showConfirmOnPanel) {
        return (this.renderConfirm(userDialogContent, currentLicense, selectedTab, showLoadingConfirmation, errorUpdating))
    } else {
      return (
        <>
        {showConfirmation && showConfirmOnPanel &&
          <Dialog open={true}>
            <DialogSurface className={"styles.managementDialog"}>
              <DialogBody>
                <DialogTitle>
                  {"Confirmar cambios"}
                </DialogTitle>
                {this.renderConfirm(userDialogContent, currentLicense, selectedTab, showLoadingConfirmation, errorUpdating)}
              </DialogBody>
            </DialogSurface>
          </Dialog >
        }
          <DialogContent className={styles.DepartmentDetails}>
            {/* Pestañas */}
            {showMode===componentTypeMode.All &&
              <TabList
                className={styles.tabList}
                size="small"
                selectedValue={selectedTab}
                onTabSelect={(event, data) => this.setState({ selectedTab: data.value as string })}
              >
                <Tab
                  className={styles.tab}
                  value="prodpiedades"
                  icon={<Settings20Regular />}
                >
                  {strings.properties}
                </Tab>
                {((editSelectedItem && editSelectedItem.itemType === SearchType.User) || (organDialogContent && !organDialogContent.IsEXTERNAL)) &&
                  <Tab
                    className={styles.tab}
                    value="Listado"
                    icon={editSelectedItem?.itemType === SearchType.User ? <ListBar20Regular /> : <PeopleTeam20Regular />}
                  >
                    {listingTabName}
                  </Tab>
                }
                {(editSelectedItem && editSelectedItem.itemType === SearchType.User) &&
                  <Tab
                    className={styles.tab}
                    value="Licencias"
                    icon={<ContactCardRibbon20Regular/>}
                  >
                    {strings.LicenseTitle}
                  </Tab>
                }
                {/*editSelectedItem?.itemType === SearchType.Organ &&
                <Tab
                  className={styles.tab}
                  value="Plantillas"
                  icon={<Folder20Regular />}
                >
                  {strings.Templates}
                </Tab>
              */}
              </TabList>
            }
            {/* Contenido */}
            <div>
              {
                (selectedTab === "prodpiedades") &&
                <div>
                  {
                    /* prodpiedades de un órgano */
                    (editSelectedItem?.itemType === SearchType.Organ && organDialogContent) &&
                    <>
                      <div className={styles.gridContainerSingle}>
                        {/* Nombre del órgano */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.DepartmentName}
                          </Label>
                          <Input
                            className={formErrors && formErrors.Department ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.NombreDepartment}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "NombreDepartment", data.value)}
                          />
                          {/*  
                            <Dropdown
                              placeholder={strings.DepartmentNamePlaceholder}
                              className={styles.dropDownValue}
                              defaultValue={organDialogContent.Department ? this.props.organDictionary[organDialogContent.Department] : ""}
                              onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Department", data.optionValue)}
                            >
                              {
                                organFilterOptions.map(option =>
                                  <Option
                                    key={option.key}
                                    value={option.key}
                                  >
                                    {option.text}
                                  </Option>
                                )
                              }
                            </Dropdown>
                          */}
                        </div>
                        {formErrors && formErrors.Department &&
                          <div className={styles.propertyRow}>
                            <Label className={styles.labelproperty}>
                            </Label>
                            <Label
                              className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                            >
                              <Info12Filled className={styles.errorIcon} />{formErrors.Department}
                            </Label>
                          </div>
                        }
                        {/* Descripción */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Description}
                          </Label>
                          <Textarea
                            className={styles.textBoxValue}
                            value={organDialogContent.Description}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Description", data.value)}
                          />
                        </div>
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Tipo de órgano */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.OrganType}
                          </Label>
                          <Dropdown
                            className={styles.dropDownValue}
                            value={organDialogContent.TipoDepartment ? this.props.organTypeDictionary[organDialogContent.TipoDepartment] : ""}
                            disabled={!editSelectedItem?.isNewItem && !organDialogContent?.IsEXTERNAL}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "TipoDepartment", data.optionValue)}
                          >
                            {
                              organTypeFilterOptions.map(option =>
                                <Option
                                  key={option.key}
                                  value={option.key}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {/* Area Sectorial*/}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.BusinessArea}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.BusinessArea ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.BusinessArea ? BusinessAreaOptions.find(opt => opt.key === organDialogContent.BusinessArea)?.text : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "BusinessArea", data.optionValue)}
                          >
                            {
                              BusinessAreaOptions.map(option =>
                                <Option
                                  key={option.key}
                                  value={option.key}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {formErrors && formErrors.BusinessArea &&
                          <>
                            <div className={styles.propertyRow}>
                            </div>
                            <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              <Label
                                className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                              >
                                <Info12Filled className={styles.errorIcon} />{formErrors.BusinessArea}
                              </Label>
                            </div>
                          </>
                        }
                        {/* Nomenclatura */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} >
                            {strings.Nomenclature}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={organDialogContent.BodyId}
                            disabled={!editSelectedItem?.isNewItem && !organDialogContent?.IsEXTERNAL}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "BodyId", data.value)}
                          />
                        </div>
                        {/* Division */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Ministry}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.Division ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.Division ? this.props.ministryDictionary[organDialogContent.Division] : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Division", data.optionValue)}
                          >
                            {
                              ministryFilterOptions.map(option =>
                                <Option
                                  key={option.key}
                                  value={option.key}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {formErrors && formErrors.Division &&
                          <>
                            <div className={styles.propertyRow}>
                            </div>
                            <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              <Label
                                className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                              >
                                <Info12Filled className={styles.errorIcon} />{formErrors.Division}
                              </Label>
                            </div>
                          </>
                        }
                        {/* Abreviatura */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Abbreviation}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={organDialogContent.Abreviatura}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Abreviatura", data.value)}
                          />
                        </div>
                        {/* Secretaría */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Secretaria}
                          </Label>
                          <Dropdown
                            className={styles.dropDownValue}
                            value={organDialogContent.Secretaria ? secretariaOptions.find(opt => opt.key === organDialogContent.Secretaria)?.text : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Secretaria", data.optionValue)}
                          >
                            {
                              secretariaOptions.map(option =>
                                <Option
                                  key={option.key}
                                  value={option.key}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {/* Intersectorial */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Intersectorial}
                          </Label>
                          <Switch
                            //TODO REVISAR CAMPO
                            className={styles.dropDownValue}
                            checked={organDialogContent.Intersectorial}
                            label={organDialogContent.Intersectorial ? strings.IntersectorialYes : strings.IntersectorialNo}
                            //onChange={(ev, data) => { this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Intersectorial: data.checked } } }) }}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Intersectorial", data.checked)}
                          />
                        </div>
                        {/* Period */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.Legislature}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.Period ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.Period ? legistatureOptions.find(opt => opt.key === organDialogContent.Period)?.text : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Period", data.optionValue)}
                          >
                            {
                              legistatureOptions.map(option =>
                                <Option
                                  key={option.key}
                                  value={option.key}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {formErrors && formErrors.Period &&
                          <>
                            <div className={styles.propertyRow}>
                            </div>
                            <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              <Label
                                className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                              >
                                <Info12Filled className={styles.errorIcon} />{formErrors.Period}
                              </Label>
                            </div>
                          </>
                        }
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Fecha constitución */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.CreationDate}
                          </Label>
                          <Input
                            key='StartDate'
                            value={organDialogContent?.FechaConstitucion && format(new Date(organDialogContent.FechaConstitucion), 'yyyy-MM-dd')}
                            type='date'
                            onChange={async (data): promise<void> =>
                              // await setFieldValue("EndDate", new Date(data.target.value).toISOString());
                              //this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, IncorporationDate: new Date(data.target.value).toISOString() } } })
                              this.onUpdateproperty(editSelectedItem.itemType, "FechaConstitucion", new Date(data.target.value).toISOString())
                            }
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                          />
                        </div>
                        {/* Fecha extinción */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.ExpirationDate}
                          </Label>
                          <Input
                            key='EndDate'
                            value={organDialogContent?.FechaExtincion && format(new Date(organDialogContent.FechaExtincion), 'yyyy-MM-dd')}
                            type='date'
                            onChange={async (data): promise<void> => {
                              this.onUpdateproperty(editSelectedItem.itemType, "FechaExtincion", new Date(data.target.value).toISOString())
                            }}
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                            disabled={!editSelectedItem?.isNewItem && !organDialogContent.IsEXTERNAL}
                          />
                        </div>
                        {/* SIA */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.SIA}
                          </Label>
                          <Input
                            type='text'
                            className={formErrors && formErrors.SIA ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.SIA}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "SIA", data.value)}
                          />
                        </div>
                        {/* DIR3 */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.DIR3}
                          </Label>
                          <Input
                            type='text'
                            className={formErrors && formErrors.dir ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.dir}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "dir", data.value)}
                          />
                        </div>
                        {formErrors && (formErrors.SIA || formErrors.dir) &&
                          <>
                            {formErrors.SIA ?
                              <div className={styles.propertyRow}>
                                <Label className={styles.labelproperty}>
                                </Label>
                                <Label
                                  className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                >
                                  <Info12Filled className={styles.errorIcon} />{formErrors.SIA}
                                </Label>
                              </div>
                              :
                              <div className={styles.propertyRow}>
                              </div>
                            }
                            {formErrors.dir &&
                              <div className={styles.propertyRow}>
                                <Label className={styles.labelproperty}>
                                </Label>
                                <Label
                                  className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                >
                                  <Info12Filled className={styles.errorIcon} />{formErrors.dir}
                                </Label>
                              </div>
                            }
                          </>
                        }
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Días para aprodbar el Minutes */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.DaysToApprodve}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            type='number'
                            value={organDialogContent?.DiasAprodbacionMinutes ? organDialogContent.DiasAprodbacionMinutes.toString() : undefined}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "DiasAprodbacionMinutes", +data.value)}
                          />
                        </div>
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerSingle}>
                        {/* Observaciones */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Observations}
                          </Label>
                          <Textarea
                            className={styles.textBoxValue}
                            value={organDialogContent.Observaciones}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Observaciones", data.value)}
                          />
                        </div>
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      {/* Activo */}
                      {!editSelectedItem?.isNewItem &&
                        <div className={styles.propertyRowActive}>
                          <Label className={styles.labelproperty}>
                            {strings.ActiveStatus}
                          </Label>
                          <Label
                            className={organDialogContent.Activo ? styles.activeStatusValue : styles.inActiveStatusValue}
                          >
                            {organDialogContent.Activo ? strings.ActiveYes : strings.ActiveNo}
                          </Label>
                        </div>
                      }
                    </>
                  }
                  {
                    /* prodpiedades de un usuario */
                    (editSelectedItem?.itemType === SearchType.User && userDialogContent) &&
                    <>
                      <div className={styles.gridContainerSingle}>
                        {/* Nombre del usuario */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.DisplayName}
                          </Label>
                          <Input
                            className={formErrors && formErrors.DisplayName ? `${styles.textBoxValue} ${styles.errorField}` : styles.textBoxValue}
                            value={userDialogContent.DisplayName}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "DisplayName", data.value)}
                          />
                        </div>
                        {formErrors && formErrors.DisplayName &&
                          <div className={styles.propertyRow}>
                            <Label className={styles.labelproperty}>
                            </Label>
                            <Label
                              className={`${styles.textBoxValue} ${styles.errorMsgForm}`}
                            >
                              <Info12Filled className={styles.errorIcon} />{formErrors.DisplayName}
                            </Label>
                          </div>
                        }
                      </div>
                      <div className={styles.gridContainerDual}>
                        {/* Nombre del usuario */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.Name}
                          </Label>
                          <Input
                            className={formErrors && formErrors.FirstName ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={userDialogContent.FirstName}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "FirstName", data.value)}
                          />
                        </div>
                        {/* Apellidos del usuario */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.SurName}
                          </Label>
                          <Input
                            className={formErrors && formErrors.LastName ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={userDialogContent.LastName}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "LastName", data.value)}
                          />
                        </div>
                        {formErrors && (formErrors.FirstName || formErrors.LastName) &&
                          <>
                            <div className={styles.propertyRow}>
                              {formErrors.FirstName &&
                                <> <Label className={styles.labelproperty}>
                                </Label>
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.FirstName}
                                  </Label></>
                              }
                            </div>
                            <div className={styles.propertyRow}>
                              {formErrors.LastName &&
                                <> <Label className={styles.labelproperty}>
                                </Label>
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.LastName}
                                  </Label></>
                              }
                            </div>
                          </>
                        }
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Cargo */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Job}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={userDialogContent.JobTitle}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "JobTitle", data.value)}
                          />
                        </div>
                        {/* Comunidad autónoma */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.AutonomousCommunity}
                          </Label>
                          <Dropdown
                            className={styles.dropDownValue}
                            value={userDialogContent.Office}
                            onOptionSelect={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Office", data.optionValue)}
                          >
                            {
                              ccaaFilterOptions.map(option =>
                                <Option
                                  key={option.text}
                                  value={option.text}
                                >
                                  {option.text}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                      </div>
                      <div className={styles.gridContainerDual}>
                        {/* Division / Gob autonomico */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.CompanyNameLabel}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={userDialogContent.CompanyName}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "CompanyName", data.value)}
                          />
                        </div>
                        {/* Unidad / consejería */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.DepartmentLabel}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={userDialogContent.Department}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Department", data.value)}

                          />
                        </div>
                      </div>
                      <div className={styles.gridContainerDual}>
                        {/* tipo empleado */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.EmployeeTypeLabel}
                          </Label>
                          <Input
                            className={styles.dropDownValue}
                            value={userDialogContent.EmployeeType}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "EmployeeType", data.value)}
                          />
                        </div>
                        <div className={styles.propertyRow}>
                        </div>
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Correo electrónico */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Email}
                          </Label>
                          {!editSelectedItem?.isNewItem ?
                            <Label
                              className={styles.dropDownValue}
                            //value={userDialogContent.Email}
                            // onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Email", data.value)}
                            >
                              {userDialogContent.PrincipalMail ? userDialogContent.PrincipalMail : ""}
                            </Label>
                            :
                            <Input
                              required
                              className={styles.dropDownValue}
                              value={userDialogContent.PrincipalMail}
                              onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "PrincipalMail", data.value)}
                            />
                          }
                        </div>
                        {/* Correo electrónico secundario */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.SecondaryMail}
                          </Label>
                          <Input
                            className={formErrors && formErrors.Email ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={userDialogContent.Email}
                            onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "Email", data.value)}
                          />
                        </div>
                        {formErrors && formErrors.Email &&
                          <>
                            <div className={styles.propertyRow}>
                            </div>
                            <div className={styles.propertyRow}>
                              {formErrors.Email &&
                                <> <Label className={styles.labelproperty}>
                                </Label>
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.Email}
                                  </Label></>
                              }
                            </div>
                          </>
                        }
                        </div>
                        <div className={styles.gridContainerDual}>
                          {/* Número de teléfono móvil */}
                          <div className={styles.propertyRow}>
                            <Label className={styles.labelproperty}>
                              {strings.PhoneNumber}
                            </Label>
                            <Input
                              className={formErrors && formErrors.CellPhone ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                              value={userDialogContent.CellPhone}
                              //type='number'
                              onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "CellPhone", data.value)}
                            />
                          </div>
                             <div className={styles.propertyRow}>
                            <Label className={styles.labelproperty}>
                              {strings.BusinessPhoneLabel}
                            </Label>
                            <Input
                              className={formErrors && formErrors.BusinessPhone ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                              value={userDialogContent.BusinessPhone}
                              //type='number'
                              onChange={(ev, data) => this.onUpdateproperty(editSelectedItem.itemType, "BusinessPhone", data.value)}
                            />
                          </div>
                          {formErrors && (formErrors.CellPhone || formErrors.BusinessPhone) &&
                          <>
                            <div className={styles.propertyRow}>
                              {formErrors.CellPhone &&
                                <>
                                  <Label className={styles.labelproperty}>
                                  </Label>
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.CellPhone}
                                  </Label></>
                              }
                            </div>
                            <div className={styles.propertyRow}>
                              {formErrors.BusinessPhone &&
                                <>
                                  <Label className={styles.labelproperty}>
                                  </Label>
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.BusinessPhone}
                                  </Label></>
                              }
                            </div>
                          </>
                        }                       
                      </div>
                    </>
                  }
                </div>
              }
              {
                (selectedTab === "Listado") && <>
                  <div className={styles.fullListContainer}>
                    {/*Cabecera */
                      <div className={styles.gridListHeaderContainer}>
                        <div className={styles.listRowHeader}>
                          {editSelectedItem?.itemType === SearchType.Organ ?
                            <Person20Regular className={styles.personIcon} style={{ visibility: "hidden" }} />
                            :
                            <div className={styles.needLicenseMsg} style={{ visibility: "hidden" }}>
                              <Warning20Regular className={styles.personIcon} />
                            </div>
                          }
                          <div className={styles.listItemTitle}>
                            {editSelectedItem?.itemType === SearchType.Organ ? strings.Members : strings.Organs}
                          </div>
                          <div className={styles.listItemSelectorOrgan}>
                            {strings.Role}
                          </div>
                          {/*editSelectedItem?.itemType === SearchType.Organ &&
                            <div style={{ marginLeft: 20 }} className={editSelectedItem?.itemType === SearchType.User ? styles.listItemSelector : styles.listItemSelectorOrgan}>
                              {strings.AutonomousCommunity}
                            </div>
                          */}
                          <Button size="small" icon={<Dismiss20Regular />} style={{ visibility: "hidden" }} />
                          <Button size="small" icon={<Dismiss20Regular />} style={{ visibility: "hidden" }} />
                        </div>
                        <Divider className={styles.dividerHorizontal} />
                      </div>
                    }
                    {
                      /* Usuarios de un órgano */
                      (editSelectedItem?.itemType === SearchType.Organ && organDialogContent && organDialogContent.Users && organDialogContent.Users.length > 0) &&
                      <div className={styles.gridListContainer}>
                        {
                          organDialogContent.Users.map(user => {
                            //const needLicenseForCurrentRole = user.RoleId && user.UserPrincipalName && this.onCheckRoleLicense(user.RoleId, user.UserPrincipalName);
                            const needLicenseForCurrentRole = false;
                            return (
                              <>
                                <div className={styles.listRow} key={user.UserPrincipalName}>
                                  {needLicenseForCurrentRole ?
                                    <div className={styles.needLicenseMsg}>
                                      <Tooltip content={strings.NeedLicense} relationship='label' withArrow>
                                        <Warning20Regular className={styles.personIcon} />
                                      </Tooltip>
                                    </div>
                                    :
                                    <Person20Regular className={styles.personIcon} />
                                  }
                                  <div className={styles.listItemTitle}>
                                    <div><span>{userDictionary ? userDictionary[user.UserPrincipalName.toLowerCase()]?.DisplayName : ""}</span></div>
                                    <div className={styles.listUserJob}><span>{userDictionary ? userDictionary[user.UserPrincipalName.toLowerCase()]?.JobTitle : ""}</span></div>
                                  </div>
                                  <Dropdown
                                    className={styles.listItemSelector}
                                    placeholder={strings.SelectRolePlaceholder}
                                    //TODO MAPEO DE ROLES Y USUARIOS
                                    value={user.RoleId && roleDictionary ? roleDictionary[user.RoleId] : ""}
                                    onOptionSelect={(ev, data) => this.onUpdateListItem(editSelectedItem.itemType, 'update', user.UserPrincipalName ? user.UserPrincipalName : "", data.optionValue)}
                                  >
                                    {
                                      roleFilterOptions.map(option =>
                                        <Option
                                          key={option.key}
                                          value={option.key}
                                        >
                                          {option.text}
                                        </Option>
                                      )
                                    }
                                  </Dropdown>
                                  <Dropdown
                                    disabled
                                    className={styles.listItemSelector}
                                    placeholder={strings.SelectCCAAPlaceholder}
                                    value={user.Ccaa ? user.Ccaa : ""}
                                    onOptionSelect={(ev, data) => this.onUpdateListItem(editSelectedItem.itemType, 'update', user.UserPrincipalName ? user.UserPrincipalName : "", "", data.optionValue)}
                                    style={{ display: "none" }}
                                  >
                                    {
                                      ccaaFilterOptions.map(option =>
                                        <Option
                                          key={option.text}
                                          value={option.text}
                                        >
                                          {option.text}
                                        </Option>
                                      )
                                    }
                                  </Dropdown>
                                  {isMobile ?
                                    <>
                                      <Button className={styles.removeButton} icon={<Dismiss20Regular />} onClick={() => this.onUpdateListItem(editSelectedItem.itemType, 'delete', user.UserPrincipalName ? user.UserPrincipalName : "")}>{strings.Remove}</Button>
                                      <Divider className={"styles.dividerHorizontal"} />
                                    </>
                                    :
                                    <>
                                      <Button size="small" icon={<Dismiss20Regular />} onClick={() => this.onUpdateListItem(editSelectedItem.itemType, 'delete', user.UserPrincipalName ? user.UserPrincipalName : "")} />
                                      {showCheckPermissionsModal && checkPermissionsInfo && checkPermissionsInfo.UserPrincipalName === user.UserPrincipalName ?
                                        <Popover open={showCheckPermissionsModal} withArrow>
                                          <PopoverTrigger disableButtonEnhancement>
                                            <Tooltip content={strings.CheckPermissionsTooltip} relationship='label' withArrow>
                                              <Button style={{ visibility: user?.RoleId === RoleTaxonomyIds.Invitado ? "hidden" : "visible" }} size="small" icon={<PersonKey20Regular />} onClick={() => this.onClickCheckPermissions(user.UserPrincipalName, userDictionary ? userDictionary[user.UserPrincipalName.toLowerCase()]?.DisplayName : "", organDialogContent.BodyId, user.RoleId ? user.RoleId : "")} />
                                            </Tooltip>
                                          </PopoverTrigger>
                                          <PopoverSurface tabIndex={-1} className={styles.popOverUser}>
                                            <div className={styles.popOverUserTitle}>
                                              {strings.CheckPermissionsLabel}
                                            </div>
                                            <Divider />
                                            {checkPermissionsInfo.isLoading ?
                                              <Spinner size='medium' className={styles.popOverUserLoader} />
                                              :
                                              <div>
                                                <div className={styles.popOverUserRow}>
                                                  {`${strings.UserLabel}:  `}
                                                  <Label className={styles.popOverUserRowValue}>{checkPermissionsInfo?.DisplayName}</Label>
                                                </div>
                                                <div className={styles.popOverUserRow}>
                                                  {`${strings.RoleLabel}:  `}
                                                  <Label className={styles.popOverUserRowValue}>{checkPermissionsInfo.Role && roleDictionary ? roleDictionary[checkPermissionsInfo.Role] : ""}</Label>
                                                </div>
                                                <div className={styles.popOverUserRow}>
                                                  <Label className={checkPermissionsInfo?.HasPermissions ? `${styles.popOverUserRowValue} ${styles.hasPermission}` : `${styles.popOverUserRowValue} ${styles.notPermission}`}>
                                                    {checkPermissionsInfo?.HasPermissions
                                                      ? strings.UserHasPermissionsTxt
                                                      : checkPermissionsInfo?.AssignmentStatus === 'New'
                                                        ? 'El usuario aún no dispone de los permisos en el rol indicado. La Request está pendiente de ser prodcesada'
                                                        : checkPermissionsInfo?.AssignmentStatus === 'In prodgress'
                                                          ? 'El usuario aún no dispone de los permisos en el rol indicado. La Request está siendo prodcesada'
                                                          : checkPermissionsInfo?.IsErrorPermanente
                                                            ? 'El usuario aún no dispone de los permisos en el rol indicado. Se ha prodducido un error prodcesando la Request. Quite al usuario del órgano y vuelva a añadirlo para intentar la asignación de nuevo'
                                                            : checkPermissionsInfo?.AssignmentStatus === 'Error'
                                                              ? 'El usuario aún no dispone de los permisos en el rol indicado. Se ha prodducido un error prodcesando la Request. Se volverá a intentar en unos minutos'
                                                              : strings.UserNotPermissionsTxt
                                                    }
                                                  </Label>
                                                </div>
                                                <div className={styles.popOverUserButton}>
                                                  <Button appearance="secondary" onClick={() => this.setState({ showCheckPermissionsModal: false, checkPermissionsInfo: undefined })}>{strings.Accept}</Button>
                                                </div>
                                              </div>
                                            }
                                          </PopoverSurface>
                                        </Popover>
                                        :
                                        <Tooltip content={strings.CheckPermissionsTooltip} relationship='label' withArrow>
                                          <Button style={{ visibility: user?.RoleId === RoleTaxonomyIds.Invitado ? "hidden" : "visible" }} disabled={showCheckPermissionsModal} size="small" icon={<PersonKey20Regular />} onClick={() => this.onClickCheckPermissions(user.UserPrincipalName, userDictionary ? userDictionary[user.UserPrincipalName.toLowerCase()]?.DisplayName : "", organDialogContent.BodyId, user.RoleId ? user.RoleId : "")} />
                                        </Tooltip>
                                      }
                                    </>
                                  }
                                </div>
                              </>
                            );
                          })
                        }
                      </div>
                    }
                    {editSelectedItem?.itemType === SearchType.Organ && organDialogContent &&
                      <div className={styles.gridListContainer}>
                        {
                          newValues.Items.map(user => {
                            const itemExist = user.UserId ? (organDialogContent.Users.some(u => u.UserPrincipalName === user.UserId) || newValues.Items.some(u => (u.UserId === user.UserId && u.IndexKey !== user.IndexKey))) : false;
                            //const needLicenseForCurrentRole = (user.Role && user.UserId) ? this.onCheckRoleLicense(user.Role, user.UserId) : false;
                            const needLicenseForCurrentRole = false;
                            return (
                              <><div className={itemExist ? `${styles.listRow} ${styles.listRowError}` : `${styles.listRow}`} key={user.IndexKey}>
                                {itemExist &&
                                  <div className={styles.errorItemDuplicate}>
                                    <Tooltip content={strings.UserDuplicated} relationship='label' withArrow>
                                      <ErrorCircle20Filled className={styles.personIcon} />
                                    </Tooltip>
                                  </div>
                                }
                                {!itemExist && needLicenseForCurrentRole &&
                                  <div className={styles.needLicenseMsg}>
                                    <Tooltip content={strings.NeedLicense} relationship='label' withArrow>
                                      <Warning20Regular className={styles.personIcon} />
                                    </Tooltip>
                                  </div>
                                }
                                <div className={editSelectedItem?.itemType === SearchType.User ? styles.peoplePicker : (needLicenseForCurrentRole || itemExist) ? styles.peoplePickerOrganWarning : styles.peoplePickerOrgan}>
                                  <PeoplePicker
                                    type={PersonType.person}
                                    userType={UserType.user}
                                    transitiveSearch={true}
                                    selectionChanged={async (e) => {
                                      const upn = e && e.detail && e.detail.length > 0 ? e.detail[0] as { userPrincipalName: string } : undefined;
                                      if (upn && !organDialogContent.Users.some(u => u.UserPrincipalName?.toLowerCase() === upn?.userPrincipalName?.toLowerCase())) {
                                        this.onUpdateNewListItems(editSelectedItem.itemType, 'update', upn ? upn.userPrincipalName : undefined, user.Role, user.IndexKey, userDictionary && upn ? userDictionary[upn?.userPrincipalName?.toLowerCase()]?.Office : "")
                                      } else {
                                        this.onUpdateNewListItems(editSelectedItem.itemType, 'update', upn ? upn.userPrincipalName : undefined, user.Role, user.IndexKey, userDictionary && upn ? userDictionary[upn?.userPrincipalName?.toLowerCase()]?.Office : "")
                                      }
                                    }}
                                    selectionMode='single'
                                    //groupId={"85bc6717-b518-4c97-ba11-30056b8bef61"}
                                    placeholder={strings.SelectUserPlaceholder}
                                    defaultValue={user?.UserId}
                                    selectedPeople={user?.UserId ? [{ userPrincipalName: user?.UserId, displayName: userDictionary[user?.UserId?.toLowerCase()]?.DisplayName ? userDictionary[user?.UserId?.toLowerCase()]?.DisplayName : user?.UserId }] : []}
                                    className={styles.removeECNTyInputPicker}
                                  />
                                </div>
                                <Dropdown
                                  className={styles.listItemSelector}
                                  placeholder={strings.SelectRolePlaceholder}
                                  value={user.Role && roleDictionary ? roleDictionary[user.Role] : ""}
                                  onOptionSelect={(ev, data) => this.onUpdateNewListItems(editSelectedItem.itemType, 'update', user.UserId, data.optionValue, user.IndexKey)}
                                >
                                  {
                                    roleFilterOptions.map(option =>
                                      <Option
                                        key={option.key}
                                        value={option.key}
                                      >
                                        {option.text}
                                      </Option>
                                    )
                                  }
                                </Dropdown>
                                <Dropdown
                                  className={styles.listItemSelector}
                                  placeholder={strings.SelectCCAAPlaceholder}
                                  value={user.Ccaa ? user.Ccaa : ""}
                                  onOptionSelect={(ev, data) => this.onUpdateNewListItems(editSelectedItem.itemType, 'update', user.UserId ? user.UserId.toLowerCase() : "", "", user.IndexKey, data.optionValue)}
                                  disabled
                                  style={{ display: "none" }}
                                >
                                  {
                                    ccaaFilterOptions.map(option =>
                                      <Option
                                        key={option.text}
                                        value={option.text}
                                      >
                                        {option.text}
                                      </Option>
                                    )
                                  }
                                </Dropdown>
                                {isMobile ?
                                  <>
                                    <Button className={styles.removeButton} icon={<Dismiss20Regular />} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'delete', undefined, undefined, user.IndexKey)}>{strings.Remove}</Button>
                                    <Divider className={"styles.dividerHorizontal"} />
                                  </>
                                  :
                                  <>
                                    <Button size="small" icon={<Dismiss20Regular />} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'delete', undefined, undefined, user.IndexKey)} />
                                    <Button size="small" icon={<Dismiss20Regular />} style={{ visibility: "hidden" }} />
                                  </>
                                }
                              </div>
                              </>
                            )
                          })
                        }
                      </div>
                    }
                    {
                      /* Órganos de un usuario */
                      (editSelectedItem?.itemType === SearchType.User && userDialogContent && userDialogContent.Organ && userDialogContent.Organ.length > 0) &&
                      <div className={styles.gridListContainer}>
                        {
                          userDialogContent?.Organ?.map(organ => {
                            //const needLicenseForCurrentRole = organ.RoleId && userDialogContent.UserPrincipalName && this.onCheckRoleLicense(organ.RoleId, userDialogContent.UserPrincipalName);
                            const needLicenseForCurrentRole = false;

                            //const organModified = updatedOrgans.find(or => (or.OrganId === item.OrganId && or.RoleId !== item.RoleId));
                            return (
                              <div className={styles.listRow} key={organ.OrganId}>
                                <div className={styles.needLicenseMsg} style={{ visibility: needLicenseForCurrentRole ? "visible" : "hidden" }}>
                                  <Tooltip content={strings.NeedLicense} relationship='label' withArrow>
                                    <Warning20Regular className={styles.personIcon} />
                                  </Tooltip>
                                </div>
                                <span className={styles.listItemTitle}> {organ.DepartmentName} </span>
                                <Dropdown
                                  className={styles.listItemSelector}
                                  //value={organ.RoleId}
                                  //selectedOptions={}
                                  placeholder={strings.SelectRolePlaceholder}
                                  onOptionSelect={(ev, data) => this.onUpdateListItem(editSelectedItem.itemType, 'update', organ.OrganId ? organ.OrganId : "", data.optionValue)}

                                  value={organ.RoleId && roleDictionary ? roleDictionary[organ.RoleId] : ""}
                                //onOptionSelect={(ev, data) => this.onUpdateListItem(editSelectedItem.itemType, 'update', user.UserPrincipalName ? user.UserPrincipalName : "", data.optionValue)}
                                >
                                  {
                                    roleFilterOptions.map(option =>
                                      <Option
                                        key={option.key}
                                        value={option.key}
                                      >
                                        {option.text}
                                      </Option>
                                    )
                                  }
                                </Dropdown>
                                {isMobile ?
                                  <>
                                    <Button className={styles.removeButton} icon={<Dismiss20Regular />} onClick={() => this.onUpdateListItem(editSelectedItem.itemType, 'delete', organ.OrganId ? organ.OrganId : "")}>{strings.Remove}</Button>
                                    <Divider className={"styles.dividerHorizontal"} />
                                  </>
                                  :
                                  <>
                                    <Button size="small" icon={<Dismiss20Regular />} onClick={() => this.onUpdateListItem(editSelectedItem.itemType, 'delete', organ.OrganId ? organ.OrganId : "")} />
                                    {showCheckPermissionsModal && checkPermissionsInfo && checkPermissionsInfo.BodyId === organ.OrganId ?
                                      <Popover open={showCheckPermissionsModal} withArrow>
                                        <PopoverTrigger disableButtonEnhancement>
                                          <Tooltip content={strings.CheckPermissionsTooltip} relationship='label' withArrow>
                                            <Button style={{ visibility: organ?.RoleId === RoleTaxonomyIds.Invitado ? "hidden" : "visible" }} size="small" icon={<PersonKey20Regular />} onClick={() => this.onClickCheckPermissions(userDialogContent.UserPrincipalName, userDictionary ? userDictionary[userDialogContent.UserPrincipalName.toLowerCase()]?.DisplayName : "", organ.OrganId, organ.RoleId)} />
                                          </Tooltip>
                                        </PopoverTrigger>
                                        <PopoverSurface tabIndex={-1} className={styles.popOverUser}>
                                          <div className={styles.popOverUserTitle}>
                                            {strings.CheckPermissionsLabel}
                                          </div>
                                          <Divider />
                                          {checkPermissionsInfo.isLoading ?
                                            <Spinner size='medium' className={styles.popOverUserLoader} />
                                            :
                                            <div>
                                              <div className={styles.popOverUserRow}>
                                                {`${strings.OrganLabel}:  `}
                                                <Label className={styles.popOverUserRowValue}>{organDictionary[checkPermissionsInfo.BodyId]}</Label>
                                              </div>
                                              <div className={styles.popOverUserRow}>
                                                {`${strings.RoleLabel}:  `}
                                                <Label className={styles.popOverUserRowValue}>{checkPermissionsInfo.Role && roleDictionary ? roleDictionary[checkPermissionsInfo.Role] : ""}</Label>
                                              </div>
                                              <div className={styles.popOverUserRow}>
                                                <Label className={checkPermissionsInfo?.HasPermissions ? `${styles.popOverUserRowValue} ${styles.hasPermission}` : `${styles.popOverUserRowValue} ${styles.notPermission}`}>
                                                  {checkPermissionsInfo?.HasPermissions
                                                    ? strings.UserHasPermissionsTxt
                                                    : checkPermissionsInfo?.AssignmentStatus === 'New'
                                                      ? 'El usuario aún no dispone de los permisos en el rol indicado. La Request está pendiente de ser prodcesada'
                                                      : checkPermissionsInfo?.AssignmentStatus === 'In prodgress'
                                                        ? 'El usuario aún no dispone de los permisos en el rol indicado. La Request está siendo prodcesada'
                                                        : checkPermissionsInfo?.IsErrorPermanente
                                                          ? 'El usuario aún no dispone de los permisos en el rol indicado. Se ha prodducido un error prodcesando la Request. Quite al usuario del órgano y vuelva a añadirlo para intentar la asignación de nuevo'
                                                          : checkPermissionsInfo?.AssignmentStatus === 'Error'
                                                            ? 'El usuario aún no dispone de los permisos en el rol indicado. Se ha prodducido un error prodcesando la Request. Se volverá a intentar en unos minutos'
                                                            : strings.UserNotPermissionsTxt
                                                  }
                                                </Label>
                                              </div>
                                              <div className={styles.popOverUserButton}>
                                                <Button appearance="secondary" onClick={() => this.setState({ showCheckPermissionsModal: false, checkPermissionsInfo: undefined })}>{strings.Accept}</Button>
                                              </div>
                                            </div>
                                          }
                                        </PopoverSurface>
                                      </Popover>
                                      :
                                      <Tooltip content={strings.CheckPermissionsTooltip} relationship='label' withArrow>
                                        <Button style={{ visibility: organ?.RoleId === RoleTaxonomyIds.Invitado ? "hidden" : "visible" }} disabled={showCheckPermissionsModal} size="small" icon={<PersonKey20Regular />} onClick={() => this.onClickCheckPermissions(userDialogContent.UserPrincipalName, userDictionary ? userDictionary[userDialogContent.UserPrincipalName.toLowerCase()]?.DisplayName : "", organ.OrganId, organ.RoleId)} />
                                      </Tooltip>
                                    }
                                  </>
                                }
                              </div>
                            );
                          })
                        }
                      </div>
                    }
                    {(editSelectedItem?.itemType === SearchType.User && userDialogContent) &&
                      <div className={styles.gridListContainer}>
                        {
                          newValues.Items.map(organ => {
                            //const needLicenseForCurrentRole = (organ.Role && organ.OrganId) ? this.onCheckRoleLicense(organ.Role, userDialogContent.UserPrincipalName) : false;
                            const needLicenseForCurrentRole = false;
                            const itemExist = (organ.OrganId && userDialogContent?.Organ) ? (userDialogContent.Organ.some(o => o.OrganId === organ.OrganId) || newValues.Items.some(o => (o.OrganId === organ.OrganId && o.IndexKey !== organ.IndexKey))) : false;
                            return (
                              <div className={itemExist ? `${styles.listRow} ${styles.listRowError}` : `${styles.listRow}`} key={organ.IndexKey}>
                                {itemExist ?
                                  <div className={styles.errorItemDuplicate}>
                                    <Tooltip content={strings.OrganDuplicated} relationship='label' withArrow>
                                      <ErrorCircle20Filled className={styles.personIcon} />
                                    </Tooltip>
                                  </div>
                                  :
                                  <div className={styles.needLicenseMsg} style={{ visibility: needLicenseForCurrentRole ? "visible" : "hidden" }}>
                                    <Tooltip content={strings.NeedLicense} relationship='label' withArrow>
                                      <Warning20Regular className={styles.personIcon} />
                                    </Tooltip>
                                  </div>
                                }
                                <Combobox
                                  onOptionSelect={(_event, data): void => {
                                    this.onUpdateNewListItems(editSelectedItem.itemType, 'update', data.optionValue ? data.optionValue : undefined, organ.Role, organ.IndexKey)
                                  }}
                                  clearable
                                  placeholder={strings.SelectOrganPlaceholder}
                                  className={styles.peoplePicker}
                                  value={organ.OrganId ? organDictionary[organ.OrganId] : ""}
                                >
                                  {
                                    organFilterOptions.map((option): JSX.Element => (
                                      <Option
                                        key={option.key}
                                        value={option.key}
                                      >
                                        {option.text}
                                      </Option>
                                    ))}
                                </Combobox>
                                <Dropdown
                                  className={styles.listItemSelector}
                                  placeholder={strings.SelectRolePlaceholder}
                                  value={organ.Role && roleDictionary ? roleDictionary[organ.Role] : ""}
                                  onOptionSelect={(ev, data) => this.onUpdateNewListItems(editSelectedItem.itemType, 'update', organ.OrganId, data.optionValue, organ.IndexKey)}
                                >
                                  {
                                    roleFilterOptions.map(option =>
                                      <Option
                                        key={option.key}
                                        value={option.key}
                                      >
                                        {option.text}
                                      </Option>
                                    )
                                  }
                                </Dropdown>
                                {isMobile ?
                                  <>
                                    <Button className={styles.removeButton} icon={<Dismiss20Regular />} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'delete', undefined, undefined, organ.IndexKey)}>{strings.Remove}</Button>
                                    <Divider className={"styles.dividerHorizontal"} />
                                  </>
                                  :
                                  <>
                                    <Button size="small" icon={<Dismiss20Regular />} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'delete', undefined, undefined, organ.IndexKey)} />
                                    <Button size="small" icon={<Dismiss20Regular />} style={{ visibility: "hidden" }} />
                                  </>
                                }
                              </div>
                            )
                          })
                        }
                      </div>
                    }
                    {
                      /* Botón para añadir un usuario a un órgano */
                      editSelectedItem?.itemType === SearchType.Organ &&
                      <div className={styles.addNewItemRow}>
                        <div className={styles.divSpacing}></div>
                        <div className={styles.divSpacing}></div>
                        <div className={editSelectedItem?.itemType === SearchType.User ? styles.divButtonAddNew : styles.divButtonAddNewOrgan}>
                          <Button className={styles.addNewItemButton} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'add')}> {strings.AddNewUser} </Button>
                        </div>
                      </div>
                    }
                    {
                      /* Botón para añadir un órgano a un usuario */
                      (editSelectedItem?.itemType === SearchType.User) &&
                      <div className={styles.addNewItemRow}>
                        <div className={styles.divSpacing}></div>
                        <div className={styles.divButtonAddNew}>
                          <Button className={styles.addNewItemButton} onClick={() => this.onUpdateNewListItems(editSelectedItem.itemType, 'add')}> {strings.AddNewOrgan} </Button>
                        </div>
                      </div>
                    }
                  </div>
                  {newValues && newValues.hasError &&
                    <div className={styles.errorDuplicateMsg}>
                      {editSelectedItem?.itemType === SearchType.Organ ?
                        <label>{strings.ErrorListOrgan}</label>
                        :
                        <label>{strings.ErrorListUser}</label>
                      }
                    </div>
                  }
                </>
              }
              {(selectedTab === "Licencias" && userDialogContent) && 
                <div className={styles.licenseContainer}>
                  {licensesInformation && licensesInformation.length > 0 ?
                  <div className={styles.licenseInformation}>
                    <div className={styles.licenseHeader}>
                      {strings.LicensesStatus}
                    </div>
                    <div className={styles.licenseInfoContainer}>
                      {
                        licensesInformation.map(lic => 
                          <div className={styles.licenseInfoRow}><Label className={styles.licenseName}>{LicenseNames[lic.SkuId as keyof typeof LicenseNames]} </Label><Label>{lic.Available} {strings.AvailablesLabel} {lic.Total}</Label></div>
                        )
                      }
                    </div>
                    <Divider className={styles.divider}/>
                    {
                      console.log(userDialogContent.AssignedLicenses)
                    }
                    <div className={styles.licenseUser}>
                      <Label className={styles.licenseLabel}>{strings.License}</Label>
                      <Dropdown
                        className={styles.licenseSelector}
                        value={
                          userDialogContent && currentLicense? currentLicense!=="nolicense"? LicenseNames[currentLicense as keyof typeof LicenseNames]:"Sin licencia":userDialogContent.AssignedLicenses.length > 0 
                            ? userDialogContent.AssignedLicenses.includes(LicensesTypes.E5)
                              ? LicenseNames[LicensesTypes.E5]
                              : userDialogContent.AssignedLicenses.includes(LicensesTypes.E3)
                                ? LicenseNames[LicensesTypes.E3]
                                : userDialogContent.AssignedLicenses.includes(LicensesTypes.E5Developer)
                                  ? LicenseNames[LicensesTypes.E5Developer]
                                  : strings.NoLicense
                            : strings.NoLicense
                        }
                        selectedOptions={userDialogContent && currentLicense? currentLicense!=="nolicense"? [currentLicense]:["nolicense"]:userDialogContent.AssignedLicenses.length > 0 
                          ? userDialogContent.AssignedLicenses.includes(LicensesTypes.E5)
                            ? [LicensesTypes.E5]
                            : userDialogContent.AssignedLicenses.includes(LicensesTypes.E3)
                              ? [LicensesTypes.E3]
                              : userDialogContent.AssignedLicenses.includes(LicensesTypes.E5Developer)
                                ? [LicensesTypes.E5Developer]
                                : ["nolicense"]
                          : ["nolicense"]}
                        onOptionSelect={(ev, data) => this.setState({currentLicense:data.optionValue})}
                      >
                         <Option  
                            key={"nolicense"}
                            value={"nolicense"}
                          >
                            {strings.NoLicense}
                          </Option>
                        {
                          licensesInformation.map(option =>
                            <Option
                              key={option.SkuId}
                              value={option.SkuId}
                              disabled={!option.Available || +option.Available <=0}
                            >
                              {LicenseNames[option.SkuId as keyof typeof LicenseNames]}
                            </Option>
                          )
                        }
                      </Dropdown>
                    </div>
                  </div>
                  :
                  <div>
                    {strings.LicensesNotAvailable}
                  </div>
                  }
                </div>
              }
              {
                (selectedTab === "Plantillas") &&
                <div>
                  {
                  /*<FileList
                    className={styles.fileList}
                    pageSize={10}
                    // enableFileUpload={!locked}
                    disableOpenOnClick
                    itemClick={this.itemClick.bind(this)}
                  // fileListQuery={`/me/drives/${driveRoot?.DriveId}/items/${driveRoot?.DriveItemId}/children`}
                  // fileListQuery={`/root/drives/${driveRoot?.DriveId}/items/${driveRoot?.DriveItemId}/children`}
                  // driveId={driveRoot?.DriveId || ""}
                  // itemId={driveRoot?.DriveItemId}
                  // siteId='contoso-test.sharepoint.com,01756ac5-fecb-4c23-80ee-2f786bb7ff7b,418d1217-efa8-4c92-bd2b-fa76b074f9be'
                  // fileListQuery='/sites/contoso-test.sharepoint.com,01756ac5-fecb-4c23-80ee-2f786bb7ff7b,418d1217-efa8-4c92-bd2b-fa76b074f9be/drivehttps://graph.microsoft.com/v1.0/sites/contoso-test.sharepoint.com,01756ac5-fecb-4c23-80ee-2f786bb7ff7b,418d1217-efa8-4c92-bd2b-fa76b074f9be/drives/b!xWp1Acv-I0yA7i94a7f_excSjUGo75JMvSv6drB0-b5lBo2myhcgQKaJvnhBUa0H'
                  >
                    <FileTemplate template='file' {...this.props} onDeleteClick={this.onFileSelect.bind(this)}></FileTemplate>
                    {/*<Loading template='loading'></Loading>}
                  </FileList>
                */}
                </div>
              }
            </div>
          </DialogContent>
          <DialogActions className={styles.footerActionButtons}>
            <DialogTrigger disableButtonEnhancement>
              <Button
                appearance="secondary"
                onClick={() => this.onClickClose()}
              >
                {strings.Discard}
              </Button>
            </DialogTrigger>
            <Button
              appearance="primary"
              onClick={() => this.onClickSave()}
            >
              {strings.Save}
            </Button>
          </DialogActions>
        </>
      );
    }
  }

  private renderConfirm = (userDialogContent: IUserResult | undefined,currentLicense: string | undefined,  selectedTab: string, showLoadingConfirmation: boolean | undefined, errorUpdating?:boolean): React.ReactElement => {
    const changes = this.onGetChanges();
    return (
      <>
        {showLoadingConfirmation && !errorUpdating ?
          <div style={{ padding: 10, textAlign: "center", marginTop: 40, gridArea: "1 / 1 / 1 / 3" }}>
            <Spinner labelPosition="below" label={strings.SavingLoading} />
          </div>
          :
          <>{
            errorUpdating ?
              <DialogContent>
                <div>{strings.ErrorMsg}</div>
              </DialogContent >
              :
              <DialogContent>
                {selectedTab === "Licencias" && userDialogContent &&
                  <div>
                    {(!userDialogContent.AssignedLicenses && currentLicense && currentLicense !== "nolicense") || (currentLicense && !userDialogContent.AssignedLicenses.includes(currentLicense)) ?
                      <div>
                        {currentLicense === "nolicense" ?
                          <div>
                            <Label>{strings.DeleteLicense}</Label><Label style={{fontWeight:"bold"}}>{this.props?.managedItem?.userDialogContent?.DisplayName}</Label>
                            </div>
                        :
                          <div>
                            <Label>{strings.AssingLicense}</Label><Label style={{fontWeight:"bold"}}>{LicenseNames[currentLicense as keyof typeof LicenseNames]}</Label><Label>{strings.ForUser}</Label><Label style={{fontWeight:"bold"}}>{this.props?.managedItem?.userDialogContent?.DisplayName}</Label>
                          </div>
                        }
                      </div>
                    :
                    <div>
                      {strings.NoChanges}
                    </div>
                    }
                  </div>
                }
                {selectedTab !== "Licencias" && 
                  <>
                  {changes && changes.length > 0 ?
                    <div>
                      <MessageBar intent={"warning"} style={{ padding: "10px", marginBottom: "25px" }}>
                        <MessageBarBody>
                          <MessageBarTitle> {strings.Warning} </MessageBarTitle>
                          {strings.PermissionsTimeNote}
                        </MessageBarBody>
                      </MessageBar>
                      {strings.ChangesDone}
                      <ul>
                        {changes && changes.map(val => {
                          return (
                            <li>
                              {val.property}: <label style={{ fontWeight: 600 }}>{val.newValue ? val.newValue : selectedTab === "Listado" ? <label style={{ color: "#bc2f32" }}>{strings.RemovedListItem}</label> : strings.ECNTyValueChange}</label>
                            </li>
                          )
                        })}
                      </ul>
                      <div>{strings.AreYouSure}</div>
                    </div>
                    :
                    <div>
                      {strings.NoChanges}
                    </div>
                  }
                  </>
                }
              </DialogContent>
          }</>
        }
        <DialogActions>
          <DialogTrigger disableButtonEnhancement>
            <Button
              appearance="secondary"
              onClick={() => this.onClickCloseConfirmation()}
              disabled={showLoadingConfirmation && !errorUpdating}
            >
              {strings.Discard}
            </Button>
          </DialogTrigger>
          {!errorUpdating &&
            <Button
              appearance="primary"
              onClick={() => this.onClickSaveConfirmation()}
              disabled={
                (
                  (!changes || changes.length < 1) && selectedTab !== "Licencias"
                ) || (
                  userDialogContent &&
                  selectedTab === "Licencias" &&
                  userDialogContent.AssignedLicenses.length > 0 &&
                  !(
                    (
                      userDialogContent.AssignedLicenses.includes(LicensesTypes.E3) ||
                      userDialogContent.AssignedLicenses.includes(LicensesTypes.E5) ||
                      userDialogContent.AssignedLicenses.includes(LicensesTypes.E5Developer)
                    ) &&
                    currentLicense === "nolicense"
                  ) &&
                  !(
                    !userDialogContent.AssignedLicenses.includes(LicensesTypes.E3) &&
                    !userDialogContent.AssignedLicenses.includes(LicensesTypes.E5) &&
                    !userDialogContent.AssignedLicenses.includes(LicensesTypes.E5Developer) &&
                    currentLicense !== "nolicense"
                  )
                   || 
                      (selectedTab === "Licencias" && currentLicense && userDialogContent?.AssignedLicenses.includes(currentLicense)) || 
                      (selectedTab === "Licencias" && !currentLicense)
                ) || (
                  selectedTab === "Licencias" && currentLicense && 
                  currentLicense === "nolicense" && 
                  (!userDialogContent?.AssignedLicenses || userDialogContent?.AssignedLicenses.length < 1)
                ) || 
                showLoadingConfirmation
              }
            >
              {strings.Save}
            </Button>
          }
        </DialogActions >
      </>
    )
  }
}

export default DepartmentDetails;
