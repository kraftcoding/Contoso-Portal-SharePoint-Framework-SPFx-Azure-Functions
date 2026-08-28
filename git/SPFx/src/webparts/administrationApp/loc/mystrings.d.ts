declare interface IAdministrationAppWebPartStrings {
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
  NewUsers: string;
  NewMultipleUsers: string;

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
}

declare module 'AdministrationAppWebPartStrings' {
  const strings: IAdministrationAppWebPartStrings;
  export = strings;
}
