declare interface IMasterTablePeriodWebPartStrings {
  propertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  PaginationGroupName: string;
  PageSizeFieldLabel: string


  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;

  // Botones
  Title: string;
  Add: string;
  Edit: string;
  Delete: string;
  Refresh: string;
  Save: string;
  Cancel: string;
  Status: string;
  Active: string;
  Inactive: string;

  // Mensajes
  ConfirmDelete: string;
  IsLoadingItems: string;
  ErrorTitle: string;
  RequiredFieldsMissing: string;
  ConfirmDeleteTitle: string;
  ConfirmDeleteOne: string;
  ConfirmDeleteMany: string;
  DeleteConfirm: string;
  RequiredFieldError: string;
  NoPermissionError: string;
  NotFoundError: string;
  ReferencedError: string;
  ErrorMsg: string;
  StartDateRequiredError: string;
  DateEndGtStartError: string;
  DateOverlapError: string;

  // Columnas
  codigo: string;
  descripcion: string;
  StartDate: string;
  EndDate: string;
  activo: string;
  acciones: string;

  // Paginación
  Pagination: string;
  PagingShowing: string;
  PagingNoItems: string;
  FirstPage: string;
  testviousPage: string;
  NextPage: string;
  LastPage: string;
  PerPageFixed: string;
}

declare module 'MasterTablePeriodWebPartStrings' {
  const strings: IMasterTablePeriodWebPartStrings;
  export = strings;
}