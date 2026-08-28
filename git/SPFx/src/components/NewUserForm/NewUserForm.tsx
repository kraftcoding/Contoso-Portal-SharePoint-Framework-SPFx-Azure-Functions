import * as React from 'react';
import { INewUserFormState, INewUserFormprops } from './INewUserForm';
import styles from './NewUserForm.module.scss';
import { Button, DialogActions, DialogContent, DialogTrigger, Divider, Dropdown, Input, Label, Option, Spinner, Textarea } from '@fluentui/react-components';
import { Info12Filled } from '@fluentui/react-icons';
import { Administration, BodyUserInfo, NewUserErrorsForm, NewUserInvitation, profileInfo, UserBodyRole } from '../../webparts/administrationApp/models/AdministrationAppModels';
import strings from 'NewUserFormStrings';

class NewUserForm extends React.Component<INewUserFormprops, INewUserFormState> {
  constructor(props: INewUserFormprops | Readonly<INewUserFormprops>) {
    super(props);
    this.state = {
      userInfo:{
        UserInfo:{
          EmployeeType:"service-account@contoso.local"
        }, 
        UserInvitation:{
          InvitationMessage: Administration.invitationMessage
        }
      
      },
      showConfirmPanel:false,
      isLoadingSave: false,
      showError: false,
    };
  }

  componentDidMount(): void {
      this.setState({
        userInfo:{
        UserInfo:{
          EmployeeType:"service-account@contoso.local"
        }, 
        UserInvitation:{
          InvitationMessage: this.props.invitationMsg ? this.props.invitationMsg : Administration.invitationMessage
        }
      
      },
      })
  }
  componentDidUpdate(testvprops: Readonly<INewUserFormprops>): void {
    if (testvprops.hasErrorNewUser !== this.props.hasErrorNewUser) {
      this.setState({
        showError: this.props.hasErrorNewUser,
        isLoadingSave: false
      })
    }
  }

  private onUpdateproperty(property: string, value: string): void {
    const { userInfo } = this.state;
    let newInfoValues:BodyUserInfo = userInfo ? userInfo : {UserInfo:{}}
    newInfoValues.UserInfo = {
      ...newInfoValues.UserInfo, [property]:value
    }
    this.setState({
      userInfo: newInfoValues
    })
  }

  private onUpdateBodyRole(property: string, value: string): void{
    const { userInfo } = this.state;
    let newInfoValues:BodyUserInfo = userInfo ? userInfo : {UserInfo:{}, UserBodyRole:{}, UserInvitation:{}}
    newInfoValues.UserBodyRole = {
      ...newInfoValues.UserBodyRole, [property]:value
    }
    this.setState({
      userInfo: newInfoValues
    })
  }

  private onUpdateInvitationMessage(property: string, value: string): void{
    const { userInfo } = this.state;
    let newInfoValues:BodyUserInfo = userInfo ? userInfo : {UserInfo:{}, UserBodyRole:{}, UserInvitation:{}}
    newInfoValues.UserInvitation = {
      ...newInfoValues.UserInvitation, [property]:value
    }
    this.setState({
      userInfo: newInfoValues
    })
  }

  private onClickClose(): void{
    const { showConfirmPanel } = this.state;
    if(!showConfirmPanel){
      this.props.onClosePanel();
      this.setState({formErrors:{}, userInfo:undefined, showConfirmPanel: false, isLoadingSave: false})
    }else{
      this.setState({showConfirmPanel: false, showError: false});
    }
  }

  private onClickSave(): void{
    const { userInfo, showConfirmPanel } = this.state;
    if(!showConfirmPanel){
      const hasError = this.onCheckErrors();
      if(userInfo && !hasError){
        this.setState({showConfirmPanel:true})
      }
    }else{
      if(userInfo){
        this.setState({isLoadingSave: true});
        let userToCreate:BodyUserInfo = {...userInfo};
        if(this.props.testSelectedOrgan && userInfo?.UserBodyRole?.Role ){
          userToCreate.UserBodyRole = {
            UserUpn:"",
            Role: userInfo?.UserBodyRole?.Role,
            BodyId:this.props.testSelectedOrgan.key
          };
        }else{
          if(!userInfo.UserBodyRole || !userInfo?.UserBodyRole?.Role || !userInfo?.UserBodyRole?.BodyId){
            userToCreate.UserBodyRole = {
              UserUpn:"",
              Role:"",
              BodyId:""
            };
          }else if(userToCreate && userToCreate.UserBodyRole && userToCreate.UserBodyRole.BodyId && userToCreate.UserBodyRole.Role){
            userToCreate.UserBodyRole.UserUpn = "";
          }
        }
        userToCreate.UserInfo.PrincipalMail =  userToCreate.UserInfo.PrincipalMail?.toLowerCase();
        this.props.onCreateUser(userToCreate);
        //this.setState({formErrors:{}, userInfo:undefined, showConfirmPanel: false, isLoadingSave: false})
      }
    }
  }

  private onCheckErrors(): boolean{
    const { userInfo } = this.state;
    /*if(!userInfo || !userInfo.profileInfo){
      return true;
    }*/
    const infoValues = userInfo && userInfo.UserInfo ? userInfo.UserInfo : {};
    let hasError=false;
    let errors: NewUserErrorsForm = {}
    /*if(!infoValues?.DisplayName || !infoValues.DisplayName.trim()){
      errors.DisplayName = strings.FieldRequired;
      hasError = true;
    }
      */
    if(!infoValues?.FirstName || !infoValues.FirstName.trim()){
      errors.FirstName = strings.FieldRequired;
      hasError = true;
    }else if(infoValues.FirstName.length > 64){
      errors.FirstName = strings.MaxLong64;
      hasError = true;
    }

    if(!infoValues?.LastName || !infoValues.LastName.trim()){
      errors.LastName = strings.FieldRequired;
      hasError = true;
    }else if(infoValues.LastName.length > 64){
      errors.LastName = strings.MaxLong64;
      hasError = true;
    }

    if (!infoValues.PrincipalMail || !infoValues.PrincipalMail.trim()) {
      hasError = true;
      errors.PrincipalMail = strings.FieldRequired;
    }
    if (infoValues.PrincipalMail && !infoValues.PrincipalMail.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) {
      hasError = true;
      errors.PrincipalMail = strings.MailError;
    }
    if (infoValues.Email && !infoValues.Email.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/)) {
      hasError = true;
      errors.Email = strings.MailError;
    }
    if (infoValues.BusinessPhone && !infoValues.BusinessPhone.match(/^\d{9}$/)) {
      hasError = true;
      errors.BusinessPhone = strings.PhoneError;
    }
    if (infoValues.CellPhone && !infoValues.CellPhone.match(/^\d{9}$/)) {
      hasError = true;
      errors.CellPhone = strings.PhoneError;
    }

    if(infoValues.CompanyName && infoValues.CompanyName.length>64){
      hasError = true;
      errors.CompanyName = strings.MaxLong64
    }
    if(infoValues.Department && infoValues.Department.length>64){
      hasError = true;
      errors.Department = strings.MaxLong64
    }
    if(infoValues.JobTitle && infoValues.JobTitle.length>128){
      hasError = true;
      errors.JobTitle = strings.MaxLong128
    }
    if(infoValues.EmployeeType && infoValues.EmployeeType.length>64){
      hasError = true;
      errors.EmployeeType = strings.MaxLong64
    }
    if (userInfo && userInfo.UserInvitation && userInfo.UserInvitation.InvitationCCRecipient && !userInfo.UserInvitation.InvitationCCRecipient.match(/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/) ) {
      hasError = true;
      errors.Cco = strings.MailError;
    }
    if (userInfo && !userInfo?.UserInvitation && !userInfo?.UserInvitation?.InvitationMessage) {
      hasError = true;
      errors.Message = strings.FieldRequired;
    }

    if(this.props.testSelectedOrgan){
      if(!userInfo?.UserBodyRole || !userInfo.UserBodyRole?.Role){
        hasError = true;
        errors.Rol = strings.FieldRequired;
      }
    }
    this.setState({formErrors: errors});

    return hasError;
  }

  public render(): JSX.Element {
    const { ccaaFilterOptions, organFilterOptions, roleFilterOptions, testSelectedOrgan } = this.props;
    const { userInfo, formErrors, showConfirmPanel, isLoadingSave, showError } = this.state;
    const profileInformation:profileInfo = userInfo ? userInfo.UserInfo : {};
    const organRoleInformation:UserBodyRole = userInfo ? userInfo.UserBodyRole ? userInfo.UserBodyRole : {}:{};
    const invitationInformation: NewUserInvitation = userInfo ? userInfo.UserInvitation ?  userInfo.UserInvitation: {} : {};
    return (
      <>
      <DialogContent>
      {showConfirmPanel ? 
        <div className={styles.newUserConfirmation}>
          <div className={styles.gridContainerSingle}>
            {strings.AreYouSure}
            {isLoadingSave && !showError &&
              <Spinner/>
            }
            {showError &&
              <div>{strings.ErrorMsg}</div>
            }
          </div>
        </div>
      :
        <div className={`${styles.newUserForm}`}> 
        {/*
          <div className={styles.gridContainerSingle}>
            Nombre para mostrar
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty} required>
                {strings.DisplayName}
              </Label>
              <Input
                className={formErrors && formErrors.DisplayName ? `${styles.textBoxValue} ${styles.errorField}` : styles.textBoxValue}
                value={profileInformation.DisplayName}
                onChange={(ev, data) => this.onUpdateproperty("DisplayName", data.value)}
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
          */}
          {/* Nombre del usuario */}
          <div className={styles.gridContainerDual}>
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty} required>
                {strings.Name}
              </Label>
              <Input
                className={formErrors && formErrors.FirstName ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.FirstName}
                onChange={(ev, data) => this.onUpdateproperty("FirstName", data.value)}
              />
            </div>
             <div className={styles.propertyRow}></div>
            {formErrors && formErrors.FirstName &&
              <><div className={styles.propertyRow}>
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
               <div className={styles.propertyRow}></div>
            </>
            }
            {/* Apellidos del usuario */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty} required>
                {strings.SurName}
              </Label>
              <Input
                className={formErrors && formErrors.LastName ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.LastName}
                onChange={(ev, data) => this.onUpdateproperty("LastName", data.value)}
              />
            </div>
             <div className={styles.propertyRow}></div>
            {formErrors && formErrors.LastName &&
            <>
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
               <div className={styles.propertyRow}></div>
            </>
            }
          </div>
          <Divider className={"styles.dividerHorizontal"} />
          <div className={styles.gridContainerDual}>
            {/* Correo electrónico */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty} required>
                {strings.Email}
              </Label>
                <Input
                  required
                  className={formErrors && formErrors.PrincipalMail ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                  value={profileInformation.PrincipalMail}
                  onChange={(ev, data) => this.onUpdateproperty("PrincipalMail", data.value)}
                />
            </div>
            {/* Correo electrónico secundario */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.SecondaryMail}
              </Label>
                <Input
                  required
                  className={formErrors && formErrors.Email ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                  value={profileInformation.Email}
                  onChange={(ev, data) => this.onUpdateproperty("Email", data.value)}
                />
            </div>
            {formErrors && (formErrors.PrincipalMail || formErrors.Email)&&
              <>
                <div className={styles.propertyRow}>
                  {formErrors.PrincipalMail &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.PrincipalMail}
                      </Label></>
                  }
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
            {/* Telefono de emtestsa */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.BusinessPhoneLabel}
              </Label>
                <Input
                  className={formErrors && formErrors.BusinessPhone ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                  value={profileInformation.BusinessPhone}
                  onChange={(ev, data) => this.onUpdateproperty("BusinessPhone", data.value)}
                />
            </div>
            {/* Telefono movil*/}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.PhoneNumber}
              </Label>
                <Input
                  className={formErrors && formErrors.CellPhone ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                  value={profileInformation.CellPhone}
                  onChange={(ev, data) => this.onUpdateproperty("CellPhone", data.value)}
                />
            </div>
            {formErrors && (formErrors.BusinessPhone || formErrors.CellPhone)&&
              <>
                <div className={styles.propertyRow}>
                  {formErrors.BusinessPhone &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.BusinessPhone}
                      </Label></>
                  }
                </div>
                 <div className={styles.propertyRow}>
                  {formErrors.CellPhone &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.CellPhone}
                      </Label></>
                  }
                </div>
              </>
            }
          </div>
          <Divider className={"styles.dividerHorizontal"} />
          <div className={styles.gridContainerDual}>
            {/* Comunidad autónoma */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.AutonomousCommunity}
              </Label>
              <Dropdown
                className={styles.dropDownValue}
                value={profileInformation.Office}
                onOptionSelect={(ev, data) => this.onUpdateproperty("Office", data.optionValue ? data.optionValue : "")}
                clearable
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
            {/* Division /Gob autonomico */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.CompanyNameLabel}
              </Label>
              <Input
                className={formErrors && formErrors.CompanyName ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.CompanyName}
                onChange={(ev, data) => this.onUpdateproperty("CompanyName", data.value)}
              />
            </div>
            {formErrors && (formErrors.Office || formErrors.CompanyName)&&
              <>
                <div className={styles.propertyRow}>
                  {formErrors.Office &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.Office}
                      </Label></>
                  }
                </div>
                 <div className={styles.propertyRow}>
                  {formErrors.CompanyName &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.CompanyName}
                      </Label></>
                  }
                </div>
              </>
            }
          </div>
          <div className={styles.gridContainerDual}>
            {/* Unidad / consejeria */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.DepartmentLabel}
              </Label>
              <Input
                className={formErrors && formErrors.Department ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.Department}
                onChange={(ev, data) => this.onUpdateproperty("Department", data.value)}
              />
            </div>
            {/* Cargo */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.Job}
              </Label>
              <Input
                className={formErrors && formErrors.JobTitle ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.JobTitle}
                onChange={(ev, data) => this.onUpdateproperty("JobTitle", data.value)}
              />
            </div>
            {formErrors && (formErrors.Department || formErrors.JobTitle)&&
              <>
                <div className={styles.propertyRow}>
                  {formErrors.Department &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.Department}
                      </Label></>
                  }
                </div>
                 <div className={styles.propertyRow}>
                  {formErrors.JobTitle &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.JobTitle}
                      </Label></>
                  }
                </div>
              </>
            }
          </div>
          <div className={styles.gridContainerSingle}>
            {/* Tipo de empleado */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.EmployeeTypeLabel}
              </Label>
              <Input
                className={formErrors && formErrors.EmployeeType ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                value={profileInformation.EmployeeType}
                onChange={(ev, data) => this.onUpdateproperty("EmployeeType", data.value)}
              />
            </div>
            <div className={styles.propertyRow}></div>
            {formErrors && formErrors.EmployeeType &&
              <>
                <div className={styles.propertyRow}>
                  {formErrors.EmployeeType &&
                    <> <Label className={styles.labelproperty}>
                    </Label>
                      <Label
                        className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                      >
                        <Info12Filled className={styles.errorIcon} />{formErrors.EmployeeType}
                      </Label></>
                  }
                </div>
                 <div className={styles.propertyRow}>  
                </div>
              </>
            }
          </div>
          <Divider className={"styles.dividerHorizontal"} />
          <div className={styles.gridContainerSingle}>
            <div className={styles.gridContainerSingle}>
              {/* Órgano */}
              <div className={styles.propertyRow}>
                <Label className={styles.labelproperty} required={testSelectedOrgan !== undefined}>
                  {strings.DepartmentName}
                </Label>
                <Dropdown
                  className={styles.textBoxValue}
                  value={organRoleInformation.BodyId ? organFilterOptions.find(or => or.key === organRoleInformation.BodyId)?.text : testSelectedOrgan ? testSelectedOrgan.text : undefined}
                  onOptionSelect={(ev, data) => this.onUpdateBodyRole("BodyId", data.optionValue ? data.optionValue : "")}
                  clearable={testSelectedOrgan ? false: true}
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
              </div>
            </div>
            <div className={styles.gridContainerSingle}>
              {/* Role */}
              <div className={styles.propertyRow}>
                <Label className={styles.labelproperty} required={testSelectedOrgan !== undefined}>
                  {strings.RoleLabel}
                </Label>
                <Dropdown
                  className={styles.dropDownValue}
                  value={organRoleInformation.Role ? roleFilterOptions.find(or => or.key === organRoleInformation.Role)?.text : undefined}
                  onOptionSelect={(ev, data) => this.onUpdateBodyRole("Role", data.optionValue ? data.optionValue : "")}
                  clearable
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
              </div>
              {formErrors && formErrors.Rol &&
                <>
                  <div className={styles.propertyRow}>
                    {formErrors.Rol &&
                      <> <Label className={styles.labelproperty}>
                      </Label>
                        <Label
                          className={`${styles.dropDownValue} ${styles.errorMsgForm}`}
                        >
                          <Info12Filled className={styles.errorIcon} />{formErrors.Rol}
                        </Label></>
                    }
                  </div>
                </>
              } 
            </div>
          </div>
          <Divider className={"styles.dividerHorizontal"} />
          <div className={styles.gridContainerSingle}>
            {/* CCO */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty}>
                {strings.InvitationCCO}
              </Label>
                <Input
                  className={formErrors && formErrors.Cco ? `${styles.dropDownValue} ${styles.errorField}` : styles.dropDownValue}
                  value={invitationInformation.InvitationCCRecipient}
                  onChange={(ev, data) => this.onUpdateInvitationMessage("InvitationCCRecipient", data.value)}
                />
            </div>
          </div>
          <div className={styles.gridContainerSingle}>
            {/* invitation message */}
            <div className={styles.propertyRow}>
              <Label className={styles.labelproperty} required>
                {strings.InvitationMessage}
              </Label>
                <Textarea 
                  className={formErrors && formErrors.Message ? `${styles.dropDownValue} ${styles.errorField} ${styles.textArea}` : `${styles.dropDownValue} ${styles.textArea}`}
                  value={invitationInformation.InvitationMessage}
                  onChange={(ev, data) => {console.log(ev);this.onUpdateInvitationMessage("InvitationMessage", data.value)}}
                  resize='both'            
                />
            </div>
          </div>
        </div>
        }
      </DialogContent>
      <DialogActions>
        <DialogTrigger disableButtonEnhancement>
          <Button
            appearance="secondary"
            onClick={() => this.onClickClose()}
            disabled={isLoadingSave}
          >
            {strings.Discard}
          </Button>
        </DialogTrigger>
          <Button
            appearance="primary"
            onClick={() => this.onClickSave()}
            disabled={isLoadingSave||showError||!invitationInformation?.InvitationMessage}
          >
            {strings.Save}
          </Button>
      </DialogActions >
      </>
    );
  }
}

export default NewUserForm;
