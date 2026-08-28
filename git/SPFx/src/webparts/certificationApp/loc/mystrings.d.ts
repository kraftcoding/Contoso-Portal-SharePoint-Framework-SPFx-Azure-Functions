declare interface ICertificationAppWebPartStrings {
  //Pane
  propertyPaneDescription: string;
  BasicGroupName: string;
  NoResultsFoundPicker:string;
  LoadingPicker:string;
  TitlePane:string;
  ComponentTypePane:string;
  ItemsPerPage:string;
  ShowEXTERNAL:string;

  //AdministrationApp
  ManageOrgan:string;
  ManageUser:string;
  LoadingLabel:string;
  LoadingError:string;
  NewOrgan:string;
  NewUser:string;

  //SearchBar
  SearchTypePlaceholder:string;
  OrgansType:string;
  UsersType:string;
  SearchPlaceholder:string;
  SearchLabel:string;
  CleanSearch:string;

  //FilterZone
  SelectValue:string;
  OrganTypeFilter:string;
  OrganTypeFilterPlaceholder:string;
  UserFilter:string;
  UserFilterPlaceholder:string;
  OrganFilter:string;
  OrganFilterPlaceholder:string;
  CCAAFilter:string;
  CCAAFilterPlaceholder:string;
  RoleFilter:string;
  RoleFilterPlaceholder:string;
  LicenseFilter:string;
  LicenseFilterPlaceholder:string;
  LicenseYes:string;
  LicenseNo:string;
  RemoveFilter:string;
  ApplyFilter:string;
  BusinessAreaFilter:string;
  BusinessAreaPlaceholder:string;

  //ResultZone
  NameHeader:string;
  NomenclaturaHeader:string;
  TypeHeader:string;
  MailHeader:string;
  OrgansHeader:string;
  ManageLabel:string;
  NoResults:string;
  PageNumberTxt:string;
  FromTotalPageTxt:string;
  CertificateLabel:string;

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
  DepartmentAdscripcion: string;
  SelectOrgan:string;
  CreateNewOrgan:string;
  Inscrito:string;
  YesLabel:string;
  NoLabel:string;
  //prodPIEDADES USUARIO
  Name:string;
  SurName:string;
  DisplayName:string;
  Email:string;
  SecondaryMail:string;
  PhoneNumber:string;
  Job:string;
  AutonomousCommunity:string;

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
  DuplicatedCode: string;
  WarningNameChange: string;

  RemovedListItem:string;
  ErrorListOrgan:string;
  ErrorListUser:string;
  FechaCreacion: string;
  DesdeFilter: string;
  HastaFilter: string;
  DocumentsTooltip: string;

  ErrorMsg:string;
  Save:string;
  Discard:string;

  UploadDocument: string;
  FileRequired: string;
  TipoDepartmentRequired: string;
  TipoDocumentRequired: string;
  DepartmentRequired: string;
  SelectFolderRequired: string;
  NewFolderRequired: string;
  FolderDuplicated: string;
  FileLabel: string;
  DocumentTypeLabel: string;
  DocumentTypePlaceHolder: string;
  FechaMeetingLabel: string;
  CarpetaReunionLabel: string;
  SelectFolderPlaceholder: string;
  CreateNewFolder: string;
  NewFolderName: string;
  SavedOk: string;
  SavingLoad: string;
  Close: string;
  Cancel: string;
}

declare module 'CertificationAppWebPartStrings' {
  const strings: ICertificationAppWebPartStrings;
  export = strings;
}
