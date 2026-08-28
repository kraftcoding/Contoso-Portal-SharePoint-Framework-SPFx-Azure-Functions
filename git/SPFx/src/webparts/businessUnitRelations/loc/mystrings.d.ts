declare interface IBusinessAreaRelationsWebPartStrings {
  propertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  IsLoadingItems: string;

  //Columns
  BusinessAreaLabel: string;
  DivisionLabel: string;
  StartDateLabel: string;
  EndDateLabel: string;
  ReviewErrorLabel: string;

  //Filters
  BusinessAreaPlaceholder: string;
  DivisionPlaceholder: string;
  ActualesLabel: string;
  HistoricoLabel: string;

  //Manage
  DiscardLabel: string;
  SaveLabel: string;
  DialogTitle: string;
  ErrorItems: string;
  ConfirmLabel: string;
  AcceptLabel: string;
  CancelLabel: string;
  ErrorUpdating: string;
  NewElementLabel:string;
  ErrorNewItem:string;
  ECNTyValue:string;
  ReviewPendingChanges:string;
}

declare module 'BusinessAreaRelationsWebPartStrings' {
  const strings: IBusinessAreaRelationsWebPartStrings;
  export = strings;
}
