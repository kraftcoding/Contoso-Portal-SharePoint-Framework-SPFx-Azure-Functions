declare interface IDepartmentDetailsStrings {
  Users:string;
  Organs:string;
  properties:string;
  List:string;
  Templates:string;

  //prodPIEDADES Department
  DepartmentName:string;
  DepartmentNamePlaceholder:string;
  OrganType:string;
  Ministry:string;
  Nomenclature:string;
  Secretaria:string;
  Intersectorial:string;
  Subject:string;
  BusinessArea:string;
  Abbreviation:string;
  GeneralPlaceholder:string;
  Legislature:string;
  CreationDate:string;
  ExpirationDate:string;
  DaysToApprodve:string;
  SIA:string;
  DIR3:string;
  Description:string;
  Observations:string;
  ActiveStatus:string;
  ActiveYes:string;
  ActiveNo:string;
  IntersectorialYes:string;
  IntersectorialNo:string;
  DepartmentNameExist:string;
  CodigoDepartment:string;
  FechaInscripcion:string;
  StatusDepartment:string;

  SelectOrgan:string;
  CreateNewOrgan:string;
      
  //prodPIEDADES USUARIO
  Name:string;
  SurName:string;
  DisplayName:string;
  Email:string;
  SecondaryMail:string;
  PhoneNumber:string;
  Job:string;
  AutonomousCommunity:string;
  CompanyNameLabel:string;
  DepartmentLabel:string;
  EmployeeTypeLabel:string;
  BusinessPhoneLabel:string;
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
  NeedLicense:string;
  UserDuplicated:string;
  OrganDuplicated:string;
  CheckPermissionsTooltip:string;
  UserLabel:string;
  OrganLabel:string;
  RoleLabel:string;
  CheckPermissionsLabel:string;
  UserHasPermissionsTxt:string;
  UserNotPermissionsTxt:string;
  Accept:string;
  PermissionsTimeNote:string;
  Warning: string;
  WarningMessage: string;

  RemovedListItem:string;
  ErrorListOrgan:string;
  ErrorListUser:string;

  ErrorMsg:string;
  Save:string;
  Discard:string;

  //Licenses
  LicenseTitle: string;
  LicensesStatus:string;
  AvailablesLabel:string;
  License:string;
  NoLicense:string;
  DeleteLicense: string;
  AssingLicense: string;
  ForUser: string;
  LicensesNotAvailable: string;

}

declare module 'DepartmentDetailsStrings' {
  const strings: IDepartmentDetailsStrings;
  export = strings;
}