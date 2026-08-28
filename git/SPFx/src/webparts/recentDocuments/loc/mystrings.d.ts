declare interface IRecentDocumentsWebPartStrings {
  Settingproperties: string;
  Title: string;
  DocsToShow: string;
  DocsToShowRequired: string;
  MoreLink: string;
  MoreLinkLabel: string;
  LoadingRecentDocuments: string;
  NoRecentDocumentsAvailable: string;
  ViewMoreRecentDocuments: string;
}

declare module 'RecentDocumentsWebPartStrings' {
  const strings: IRecentDocumentsWebPartStrings;
  export = strings;
}