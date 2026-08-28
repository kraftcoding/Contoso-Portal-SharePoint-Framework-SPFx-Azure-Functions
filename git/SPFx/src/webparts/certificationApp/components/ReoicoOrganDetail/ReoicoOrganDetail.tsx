import * as React from 'react';
import { IEXTERNALDepartmentDetailstate } from './IEXTERNALOrganDetail';
import { IEXTERNALOrganDetailprops } from './IEXTERNALOrganDetail';
import styles from './EXTERNALOrganDetail.module.scss';
import { IErrorsForm, IObjectDifferences, IOrganEXTERNALResult, OrganTaxonomyIds } from '../../models/CertificationAppModels';
import {
  Button,
  DialogActions,
  DialogTrigger,
  Divider,
  Dropdown,
  Input,
  Label,
  Tab,
  TabList,
  Textarea,
  Option,
  DialogContent,
  Spinner,
  //Dialog,
  //DialogBody,
  //DialogTitle,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
  Switch,
  Combobox
} from '@fluentui/react-components';
import {
  Info12Filled,
  //Folder20Regular,
  Settings20Regular,
} from '@fluentui/react-icons';
import _ from 'lodash';
import format from 'date-fns/format';
import strings from 'CertificationAppWebPartStrings';
import { RelationAreaFields } from '../../../BusinessAreaRelations/components/BusinessAreaRelationsModels';

class EXTERNALOrganDetail extends React.Component<IEXTERNALOrganDetailprops, IEXTERNALDepartmentDetailstate> {
  constructor(props: IEXTERNALOrganDetailprops) {
    super(props);
    this.state = {
      selectedTab: "prodpiedades",
      editSelectedItem: this.props.managedItem,
      showConfirmation: false,
      showLoadingConfirmation: false,
      formErrors: {},
    };
  }

  componentDidUpdate(testvprops: Readonly<IEXTERNALOrganDetailprops>): void {
    if (!_.isEqual(testvprops.managedItem, this.props.managedItem)) {
      this.setState({
        editSelectedItem: this.props.managedItem,   
        selectedTab: "prodpiedades",
        showConfirmation: false,
        showLoadingConfirmation: false,
        formErrors: {},
        //showErrorUpdate: false
      });
    }
  }

  private onClickClose(): void {
    const { managedItem, onManageItem } = this.props;
    onManageItem(managedItem.itemId, "Close")
  }
  
  private onClickSave(): void {
    const { selectedTab } = this.state;
    if (selectedTab === "prodpiedades") {
      const hasErrors = this.onCheckValuesForm();
      console.log(hasErrors);
      if (!hasErrors) {
        this.setState({ showConfirmation: true });
      }
    } 
  }
 
  private onClickSaveConfirmation(): void {
    const { managedItem, onManageItem } = this.props;
    const { editSelectedItem, selectedTab } = this.state;
    this.setState({ showLoadingConfirmation: true })
    if (selectedTab === "prodpiedades") {
      let organUpdatedItem: IOrganEXTERNALResult = editSelectedItem?.organDialogContent;
      if (organUpdatedItem.NombreDepartment === managedItem?.organDialogContent?.NombreDepartment) {
        organUpdatedItem.NombreDepartment = ""
      }
      if(editSelectedItem.isNewItem){
        onManageItem(managedItem.itemId, "CreateNew", organUpdatedItem);
      }else{
        onManageItem(managedItem.itemId, "Updateproperties", organUpdatedItem);
      }
      
    }
  }

  private onClickCloseConfirmation(): void {
    if (this.props.errorUpdating) {
      const { managedItem, onManageItem } = this.props;
      onManageItem(managedItem.itemId, "Close")
    }
    this.setState({ showConfirmation: false });
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
      const currentItemContent =currentItem?.organDialogContent as IOrganEXTERNALResult;
      const updatedItem = this.state.editSelectedItem?.organDialogContent as IOrganEXTERNALResult;
      if (currentItemContent.NombreDepartment !== updatedItem.NombreDepartment) {
        cambios.push({ property: strings.DepartmentName, oldValue: this.props.organDictionary[currentItemContent.Department], newValue: updatedItem.NombreDepartment ? updatedItem.NombreDepartment : "" })
      }
      if (currentItemContent.TipoDepartmentEXTERNAL !== updatedItem.TipoDepartmentEXTERNAL) {
        cambios.push({ property: strings.OrganType, oldValue: this.props.organTypeDictionary[currentItemContent.TipoDepartmentEXTERNAL], newValue: this.props.organTypeDictionary[updatedItem.TipoDepartmentEXTERNAL] })
      }
      if (currentItemContent.BusinessArea !== updatedItem.BusinessArea) {
        const currentObject = currentItemContent.BusinessArea && this.props.BusinessAreaOptions.find(opt => opt.key === currentItemContent.BusinessArea) ? this.props.BusinessAreaOptions.find(opt => opt.key === currentItemContent.BusinessArea)?.text : "";
        const newObject = updatedItem.BusinessArea && this.props.BusinessAreaOptions.find(opt => opt.key === updatedItem.BusinessArea) ? this.props.BusinessAreaOptions.find(opt => opt.key === updatedItem.BusinessArea)?.text : "";

        cambios.push({ property: strings.BusinessArea, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
      }
      /*
      if (currentItemContent.Division !== updatedItem.Division) {
        cambios.push({ property: strings.Ministry, oldValue: this.props.ministryDictionary[currentItemContent.Division], newValue: this.props.ministryDictionary[updatedItem.Division] })
      }
        */
      if (currentItemContent.Abreviatura !== updatedItem.Abreviatura) {
        cambios.push({ property: strings.Abbreviation, oldValue: currentItemContent.Abreviatura, newValue: updatedItem.Abreviatura })
      }
      if (currentItemContent.CodigoDepartment !== updatedItem.CodigoDepartment) {
        cambios.push({ property: strings.CodigoDepartment, oldValue: currentItemContent.CodigoDepartment, newValue: updatedItem.CodigoDepartment })
      }

      if (currentItemContent?.StatusDepartment?.toString() !== updatedItem.StatusDepartment?.toString()) {
        console.log(currentItemContent.StatusDepartment);
        const currentObject = currentItemContent.StatusDepartment && this.props.StatusOptions.some(opt => opt.Id.toString() === currentItemContent.StatusDepartment.toString()) ? this.props.StatusOptions.find(opt => opt.Id.toString() === currentItemContent.StatusDepartment.toString())?.descripcion : "";
        const newObject = updatedItem.StatusDepartment && this.props.StatusOptions.some(opt => opt.Id.toString() === updatedItem.StatusDepartment.toString()) ? this.props.StatusOptions.find(opt => opt.Id.toString() === updatedItem.StatusDepartment.toString())?.descripcion : "";
        cambios.push({ property: strings.StatusDepartment, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
      }
      /*
      if (currentItemContent.Secretaria !== updatedItem.Secretaria) {
        const currentObject = currentItemContent.Secretaria && this.props.secretariaOptions.find(opt => opt.key === currentItemContent.Secretaria) ? this.props.secretariaOptions.find(opt => opt.key === currentItemContent.Secretaria)?.text : "";
        const newObject = updatedItem.Secretaria && this.props.secretariaOptions.find(opt => opt.key === updatedItem.Secretaria) ? this.props.secretariaOptions.find(opt => opt.key === updatedItem.Secretaria)?.text : "";

        cambios.push({ property: strings.Secretaria, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
      }
      if (currentItemContent.Period !== updatedItem.Period) {
        const currentObject = currentItemContent.Period && this.props.legistatureOptions.find(opt => opt.key === currentItemContent.Period) ? this.props.legistatureOptions.find(opt => opt.key === currentItemContent.Period)?.text : "";
        const newObject = updatedItem.Period && this.props.legistatureOptions.find(opt => opt.key === updatedItem.Period) ? this.props.legistatureOptions.find(opt => opt.key === updatedItem.Period)?.text : "";

        cambios.push({ property: strings.Legislature, oldValue: currentObject ? currentObject : "", newValue: newObject ? newObject : "" })
      }
        */
      if (currentItemContent.FechaConstitucion !== updatedItem.FechaConstitucion) {
        cambios.push({ property: strings.CreationDate, oldValue: currentItemContent?.FechaConstitucion ? format(new Date(currentItemContent.FechaConstitucion), 'dd-MM-yyyy') : "", newValue: updatedItem?.FechaConstitucion ? format(new Date(updatedItem.FechaConstitucion), 'dd-MM-yyyy') : "" })
      }
      if (currentItemContent.FechaExtincion !== updatedItem.FechaExtincion) {
        cambios.push({ property: strings.ExpirationDate, oldValue: currentItemContent?.FechaExtincion ? format(new Date(currentItemContent.FechaExtincion), 'dd-MM-yyyy') : "", newValue: updatedItem?.FechaExtincion ? format(new Date(updatedItem.FechaExtincion), 'dd-MM-yyyy') : "" })
      }
      if (currentItemContent.FechaInscripcion !== updatedItem.FechaInscripcion) {
        cambios.push({ property: strings.FechaInscripcion, oldValue: currentItemContent?.FechaInscripcion ? format(new Date(currentItemContent.FechaInscripcion), 'dd-MM-yyyy') : "", newValue: updatedItem?.FechaInscripcion ? format(new Date(updatedItem.FechaInscripcion), 'dd-MM-yyyy') : "" })
      }
      if (currentItemContent.FechaCreacion !== updatedItem.FechaCreacion) {
        cambios.push({ property: strings.FechaCreacion, oldValue: currentItemContent?.FechaCreacion ? format(new Date(currentItemContent.FechaCreacion), 'dd-MM-yyyy') : "", newValue: updatedItem?.FechaCreacion ? format(new Date(updatedItem.FechaCreacion), 'dd-MM-yyyy') : "" })
      }
      if (currentItemContent.SecretariaText !== updatedItem.SecretariaText) {
        cambios.push({ property: strings.Secretaria, oldValue: currentItemContent.SecretariaText, newValue: updatedItem.SecretariaText })
      }
      if (currentItemContent.Observaciones !== updatedItem.Observaciones) {
        cambios.push({ property: strings.Observations, oldValue: currentItemContent.Observaciones, newValue: updatedItem.Observaciones })
      }
      if (currentItemContent.Activo !== updatedItem.Activo) {
        cambios.push({ property: strings.ActiveStatus, oldValue: currentItemContent.Activo ? strings.ActiveYes : strings.ActiveNo, newValue: updatedItem.Activo ? strings.ActiveYes : strings.ActiveNo })
      }

       if (currentItemContent.Inscrito !== updatedItem.Inscrito) {
        cambios.push({ property: strings.Inscrito, oldValue: currentItemContent.Inscrito ? strings.YesLabel : strings.NoLabel, newValue: updatedItem.Inscrito ? strings.YesLabel : strings.NoLabel })
      }

       if (currentItemContent.DepartmentAdscripcion !== updatedItem.DepartmentAdscripcion) {
        cambios.push({ property: strings.DepartmentAdscripcion, oldValue: currentItemContent.DepartmentAdscripcion ? this.props.organDictionary[currentItemContent.DepartmentAdscripcion]: "", newValue: updatedItem.DepartmentAdscripcion ? this.props.organDictionary[updatedItem.DepartmentAdscripcion] : "" })
      }
    }
    return cambios;
  }
  
    private onCheckValuesForm(): boolean {
      const { editSelectedItem } = this.state;
      let hasErrors = false;
      let errors: IErrorsForm = {};
  
      if (editSelectedItem) {
        const updatedItem = editSelectedItem?.organDialogContent as IOrganEXTERNALResult;
        const currentItem = this.props.managedItem;
        const currentItemContent: IOrganEXTERNALResult = currentItem?.organDialogContent;
        if (!updatedItem.NombreDepartment) {
          hasErrors = true;
          errors.Department = strings.FieldRequired
        } else {
          if (currentItem.isNewItem || (currentItemContent && currentItemContent.NombreDepartment && currentItemContent.NombreDepartment !== updatedItem.NombreDepartment)) {
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
        
        if (!updatedItem.StatusDepartment || updatedItem.StatusDepartment === OrganTaxonomyIds.NullTaxonomy) {
          hasErrors = true;
          errors.StatusDepartment = strings.FieldRequired
        }
        if (!updatedItem.TipoDepartmentEXTERNAL || updatedItem.TipoDepartmentEXTERNAL === OrganTaxonomyIds.NullTaxonomy) {
          hasErrors = true;
          errors.TipoDepartmentEXTERNAL = strings.FieldRequired
        }
        if (!updatedItem.CodigoDepartment) {
          hasErrors = true;
          errors.CodigoDepartment = strings.FieldRequired
        }else if(currentItem.isNewItem || (currentItemContent && currentItemContent.CodigoDepartment && updatedItem.CodigoDepartment.toLowerCase().trim()!==currentItemContent.CodigoDepartment.toLowerCase().trim())){
          if(this.props.allOrganCodes.some(el=> el.toLowerCase().trim() === updatedItem.CodigoDepartment.toLowerCase().trim())){
            hasErrors = true;
            errors.CodigoDepartment = strings.DuplicatedCode;
          }
        }
      }
      this.setState({ formErrors: errors });
      return hasErrors;
    }

  private onUpdateproperty( property: string, value: string | number | boolean | undefined): void {
    const { editSelectedItem } = this.state;
    const { BusinessAreaMiniesterioRelations } = this.props;
    const organDialogContent: IOrganEXTERNALResult = editSelectedItem?.organDialogContent;
    if(property === "BusinessArea"){
      
      const Division = BusinessAreaMiniesterioRelations.find(
        el => el[RelationAreaFields.BusinessArea] === value
      );
      const newMinistery = Division ? Division[RelationAreaFields.Division] : -1;
      this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value, "Division": newMinistery } } });
    }else{
      this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, [property]: value } } });
    }
   
  }

  public render(): JSX.Element {
    const { StatusOptions, errorUpdating, BusinessAreaOptions, organTypeFilterOptions } = this.props;
    const { selectedTab, editSelectedItem, showConfirmation, showLoadingConfirmation, formErrors } = this.state;

    let organDialogContent: IOrganEXTERNALResult | undefined;
    if (editSelectedItem?.organDialogContent) {
      organDialogContent = editSelectedItem?.organDialogContent as IOrganEXTERNALResult;
    } 
    if (showConfirmation) {
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
                disabled={!changes || changes.length < 1 || showLoadingConfirmation}
              >
                {strings.Save}
              </Button>
            }
          </DialogActions >
        </>
      )
    } else {

      let adscriptionOrgans :IOrganEXTERNALResult[] =[]
      if(organDialogContent && organDialogContent.TipoDepartmentEXTERNAL && this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL]){
        if(this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL].toLowerCase() === "grupo de trabajo sectorial"){
          const comisionSectorialId = organTypeFilterOptions.find(type => type.descripcion.toLowerCase() === "comision sectorial" ||type.descripcion.toLowerCase() === "comisión sectorial");
          const conferenciaSectorialId = organTypeFilterOptions.find(type => type.descripcion.toLowerCase() === "conferencia sectorial");
          if(comisionSectorialId && conferenciaSectorialId){
            adscriptionOrgans = this.props.allOrgansData.filter(org => org.TipoDepartmentEXTERNAL === comisionSectorialId.Id.toString() || org.TipoDepartmentEXTERNAL === conferenciaSectorialId.Id.toString());
            adscriptionOrgans = adscriptionOrgans.sort((a, b) => a.NombreDepartment?.localeCompare(b.NombreDepartment));
          }
        }else if(this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL].toLowerCase() === "comision sectorial" ||this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL].toLowerCase() === "comisión sectorial"){
          const conferenciaSectorialId = organTypeFilterOptions.find(type => type.descripcion.toLowerCase() === "conferencia sectorial");
          if(conferenciaSectorialId){
            adscriptionOrgans = this.props.allOrgansData.filter(org => org.TipoDepartmentEXTERNAL === conferenciaSectorialId.Id.toString());
            adscriptionOrgans = adscriptionOrgans.sort((a, b) => a.NombreDepartment?.localeCompare(b.NombreDepartment));
          }
        }
      }
      return (
        <>
          <DialogContent className={styles.DepartmentDetails}>
            {/* Pestañas */}
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
            </TabList>
            {/* Contenido */}
            <div>
              {
                (selectedTab === "prodpiedades") &&
                <div>
                  {
                    /* prodpiedades de un órgano */
                    (organDialogContent) &&
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
                            onChange={(ev, data) => this.onUpdateproperty("NombreDepartment", data.value)}
                            title={organDialogContent.NombreDepartment}
                          />
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
                        {this.props.managedItem && this.props.managedItem.organDialogContent && organDialogContent.NombreDepartment !== this.props.managedItem.organDialogContent.NombreDepartment && !this.props.managedItem.isNewItem &&
                          <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              <Label
                                className={`${styles.dropDownValue} ${styles.warningMsg}`}
                              >
                                {strings.WarningNameChange}
                              </Label>
                            
                          </div>
                        }
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* Tipo de órgano */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.OrganType}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.TipoDepartmentEXTERNAL ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.TipoDepartmentEXTERNAL ? this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL] : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty("TipoDepartmentEXTERNAL", data.optionValue)}
                            title={organDialogContent.TipoDepartmentEXTERNAL ? this.props.organTypeDictionary[organDialogContent.TipoDepartmentEXTERNAL] : ""}
                          >
                            {
                              organTypeFilterOptions.map(option =>
                                <Option
                                  key={option.Id}
                                  value={option.Id.toString()}
                                >
                                  {option.descripcion}
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
                            value={organDialogContent && organDialogContent.BusinessArea ? BusinessAreaOptions.find(opt => organDialogContent && opt.key === organDialogContent.BusinessArea)?.text : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty( "BusinessArea", data.optionValue)}
                            title={organDialogContent && organDialogContent.BusinessArea ? BusinessAreaOptions.find(opt => organDialogContent && opt.key === organDialogContent.BusinessArea)?.text : ""}
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
                        {formErrors && (formErrors.BusinessArea || formErrors.TipoDepartmentEXTERNAL) &&
                          <>
                              <div className={styles.propertyRow}>
                                <Label className={styles.labelproperty}>
                                </Label>
                                {formErrors.TipoDepartmentEXTERNAL &&
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.TipoDepartmentEXTERNAL}
                                  </Label>
                                }
                            </div>
                            <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              {formErrors.BusinessArea &&
                                <Label
                                  className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                >
                                  <Info12Filled className={styles.errorIcon} />{formErrors.BusinessArea}
                                </Label>
                              }
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
                            onChange={(ev, data) => this.onUpdateproperty("Abreviatura", data.value)}
                            //disabled={organDialogContent?.IsEXTERNAL}
                            title={organDialogContent.Abreviatura}
                          />
                        </div>
                         <div className={styles.propertyRow}></div>
                        {/* Division 
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.Ministry}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.Division ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.Division ? this.props.ministryDictionary[organDialogContent.Division] : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty("Division", data.optionValue)}
                            title={organDialogContent.Division ? this.props.ministryDictionary[organDialogContent.Division] : ""}
                          >
                            {
                              ministryFilterOptions.map(option =>
                                <Option
                                  key={option.Id}
                                  value={option.Id}
                                >
                                  {option.descripcion}
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
                        }*/}
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
                              this.onUpdateproperty("FechaConstitucion", data.target.value ? new Date(data.target.value).toISOString(): undefined)
                            }
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent?.FechaConstitucion && format(new Date(organDialogContent.FechaConstitucion), 'dd-MM-yyyy')}
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
                              this.onUpdateproperty("FechaExtincion", data.target.value ? new Date(data.target.value).toISOString(): undefined)
                            }}
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent?.FechaExtincion && format(new Date(organDialogContent.FechaExtincion), 'dd-MM-yyyy')}
                          />
                        </div>
                        {/* Fecha Inscripcion */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.FechaInscripcion}
                          </Label>
                          <Input
                            key='InscriptDate'
                            value={organDialogContent?.FechaInscripcion && format(new Date(organDialogContent.FechaInscripcion), 'yyyy-MM-dd')}
                            type='date'
                            onChange={async (data): promise<void> =>
                              // await setFieldValue("EndDate", new Date(data.target.value).toISOString());
                              //this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, IncorporationDate: new Date(data.target.value).toISOString() } } })
                              this.onUpdateproperty("FechaInscripcion", data.target.value ? new Date(data.target.value).toISOString(): undefined)
                            }
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent?.FechaInscripcion && format(new Date(organDialogContent.FechaInscripcion), 'dd-MM-yyyy')}
                          />
                        </div>
                         {/* Fecha Creacion */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.FechaCreacion}
                          </Label>
                          <Input
                            key='CreaDate'
                            value={organDialogContent?.FechaCreacion && format(new Date(organDialogContent.FechaCreacion), 'yyyy-MM-dd')}
                            type='date'
                            onChange={async (data): promise<void> =>
                              // await setFieldValue("EndDate", new Date(data.target.value).toISOString());
                              //this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, IncorporationDate: new Date(data.target.value).toISOString() } } })
                              this.onUpdateproperty("FechaCreacion", data.target.value ? new Date(data.target.value).toISOString(): undefined)
                            }
                            onKeyDown={(e): void => e.preventDefault()}
                            className={styles.dropDownValue}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent?.FechaCreacion && format(new Date(organDialogContent.FechaCreacion), 'dd-MM-yyyy')}
                          />
                        </div>
                        {/* SIA 
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
                        </div>*/}
                        {/* DIR3 
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
                        }*/}
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerDual}>
                        {/* CodigoDepartment */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.CodigoDepartment}
                          </Label>
                          <Input
                            className={formErrors && formErrors.CodigoDepartment ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent.CodigoDepartment}
                            onChange={(ev, data) => this.onUpdateproperty("CodigoDepartment", data.value)}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent.CodigoDepartment}
                          />
                        </div>
                        {/* Situación Department */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty} required>
                            {strings.StatusDepartment}
                          </Label>
                          <Dropdown
                            className={formErrors && formErrors.StatusDepartment ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                            value={organDialogContent && organDialogContent.StatusDepartment ? StatusOptions.find(opt => organDialogContent && opt.Id.toString() === organDialogContent.StatusDepartment.toString())?.descripcion : ""}
                            onOptionSelect={(ev, data) => this.onUpdateproperty("StatusDepartment", data.optionValue)}
                            title={organDialogContent && organDialogContent.StatusDepartment ? StatusOptions.find(opt => organDialogContent && opt.Id.toString() === organDialogContent.StatusDepartment.toString())?.descripcion : ""}
                            //disabled={organDialogContent?.IsEXTERNAL}
                          >
                            {
                              StatusOptions.map(option =>
                                <Option
                                  key={option.Id}
                                  value={option.Id}
                                >
                                  {option.descripcion}
                                </Option>
                              )
                            }
                          </Dropdown>
                        </div>
                        {formErrors && (formErrors.CodigoDepartment || formErrors.StatusDepartment) &&
                          <>
                              <div className={styles.propertyRow}>
                                <Label className={styles.labelproperty}>
                                </Label>
                                {formErrors.CodigoDepartment &&
                                  <Label
                                    className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                  >
                                    <Info12Filled className={styles.errorIcon} />{formErrors.CodigoDepartment}
                                  </Label>
                                }
                            </div>
                            <div className={styles.propertyRow}>
                              <Label className={styles.labelproperty}>
                              </Label>
                              {formErrors.StatusDepartment &&
                                <Label
                                  className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                                >
                                  <Info12Filled className={styles.errorIcon} />{formErrors.StatusDepartment}
                                </Label>
                              }
                            </div>
                          </>
                        }
                      
                        {/* Department Adscripcion*/}
                        {adscriptionOrgans && adscriptionOrgans.length>0 &&
                          <div className={styles.propertyRow}>
                            <Label className={styles.labelproperty}>
                              {strings.DepartmentAdscripcion}
                            </Label>
                            <Combobox
                              className={styles.dropDownValue}
                              value={organDialogContent && organDialogContent.DepartmentAdscripcion ? adscriptionOrgans.find(opt => organDialogContent && opt.Department === organDialogContent.DepartmentAdscripcion)?.NombreDepartment : ""}
                              onOptionSelect={(ev, data) => this.onUpdateproperty( "DepartmentAdscripcion", data.optionValue)}
                              title={organDialogContent && organDialogContent.DepartmentAdscripcion ? adscriptionOrgans.find(opt => organDialogContent && opt.Department === organDialogContent.DepartmentAdscripcion)?.NombreDepartment : ""}
                              clearable
                            >
                              {
                                adscriptionOrgans.map(option =>
                                  <Option
                                    key={option.CodigoDepartment}
                                    value={option.Department}
                                  >
                                    {option.NombreDepartment}
                                  </Option>
                                )
                              }
                            </Combobox>
                          </div>
                        }
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      <div className={styles.gridContainerSingle}>
                        {/* Secretaria text */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Secretaria}
                          </Label>
                          <Input
                            className={styles.textBoxValue}
                            value={organDialogContent.SecretariaText}
                            onChange={(ev, data) => this.onUpdateproperty("SecretariaText", data.value)}
                            //disabled={organDialogContent.IsEXTERNAL}
                            title={organDialogContent.SecretariaText}
                          />
                        </div>
                        {/* Observaciones */}
                        <div className={styles.propertyRow}>
                          <Label className={styles.labelproperty}>
                            {strings.Observations}
                          </Label>
                          <Textarea
                            className={styles.textBoxValue}
                            value={organDialogContent.Observaciones}
                            onChange={(ev, data) => this.onUpdateproperty("Observaciones", data.value)}
                            //disabled={organDialogContent.IsEXTERNAL}
                          />
                        </div>
                      </div>
                      <Divider className={"styles.dividerHorizontal"} />
                      {/* Activo */}
                      <div className={styles.propertyRowActive}>
                        <div className={styles.propertyRow}>
                        <Label className={styles.labelproperty}>
                          {strings.ActiveStatus}
                        </Label>
                        {/*<Label
                          className={organDialogContent.Activo ? styles.activeStatusValue : styles.inActiveStatusValue}
                        >
                          {organDialogContent.Activo ? strings.ActiveYes : strings.ActiveNo}
                        </Label>
                        */}
                        <Switch
                          //TODO REVISAR CAMPO
                          className={styles.dropDownValue}
                          checked={organDialogContent.Activo}
                          label={organDialogContent.Activo ? strings.ActiveYes : strings.ActiveNo}
                          //onChange={(ev, data) => { this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Intersectorial: data.checked } } }) }}
                          onChange={(ev, data) => this.onUpdateproperty("Activo", data.checked)}
                        /></div>
                       {/* Inscrito */}
                       <div className={styles.propertyRow}>
                        <Label className={styles.labelproperty}>
                          {strings.Inscrito}
                        </Label>
                        {/*<Label
                          className={organDialogContent.Activo ? styles.activeStatusValue : styles.inActiveStatusValue}
                        >
                          {organDialogContent.Activo ? strings.ActiveYes : strings.ActiveNo}
                        </Label>
                        */}
                        <Switch
                          //TODO REVISAR CAMPO
                          className={styles.dropDownValue}
                          checked={organDialogContent.Inscrito}
                          label={organDialogContent.Inscrito ? strings.YesLabel : strings.NoLabel}
                          //onChange={(ev, data) => { this.setState({ editSelectedItem: { ...editSelectedItem, organDialogContent: { ...organDialogContent, Intersectorial: data.checked } } }) }}
                          onChange={(ev, data) => this.onUpdateproperty("Inscrito", data.checked)}
                        /></div>
                      </div>
                    </>
                  }
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
            {
            <Button
              appearance="primary"
              onClick={() => this.onClickSave()}
            >
              {strings.Save}
            </Button>
            }
          </DialogActions>
        </>
      );
    }
  }
}

export default EXTERNALOrganDetail;
