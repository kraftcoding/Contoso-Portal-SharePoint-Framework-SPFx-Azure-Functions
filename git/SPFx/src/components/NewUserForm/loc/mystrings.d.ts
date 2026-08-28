declare interface INewUserFormStrings {
  Users:string;
  Organs:string;
  properties:string;
  List:string;
  Templates:string;

  //prodPIEDADES Department
  DepartmentName:string;
  SelectOrgan:string;
      
  //prodPIEDADES USUARIO
  Name:string;
  SurName:string;
  DisplayName:string;
  Email:string;
  SecondaryMail:string;
  PhoneNumber:string;
  Job:string;
  AutonomousCommunity:string;
  Status:string;
  CompanyNameLabel:string;
  DepartmentLabel:string;
  EmployeeTypeLabel:string;
  BusinessPhoneLabel:string;

  InvitationMessage:string;
  InvitationCCO: string;
  
  FieldRequired:string;
  PhoneError:string;
  MailError:string;

  ECNTyValueChange:string;
  ChangesDone:string;
  AreYouSure:string;
  NoChanges:string;
  SavingLoading:string;

  //LISTADO USUARIOS / Departments
  Members:string;
  Role:string;
  Remove:string;
  AddNewUser:string;
  AddNewOrgan:string;
  SelectUserPlaceholder:string;
  SelectOrganPlaceholder:string;
  SelectRolePlaceholder:string;
  SelectCCAAPlaceholder:string;
  RoleLabel:string;
  Accept:string;
  ErrorMsg:string;
  Save:string;
  Discard:string;
  Close: string;

  StatusOk:string;
  StatusError:string;
  StatusPending:string;
  AsignOrganRoleTooltip: string;
  OrganRoleMandatoryTooltip: string;
  DownloadTemplate: string;
  UsersNotprocessed: string;
  NoUsersLoaded: string;
  MandatoryFieldsWarning:string;
  MandatoryFieldOrgans:string;
  MaxLong64:string;
  MaxLong128:string;
  MandatoryLengthFields:string;
}

declare module 'NewUserFormStrings' {
  const strings: INewUserFormStrings;
  export = strings;
}