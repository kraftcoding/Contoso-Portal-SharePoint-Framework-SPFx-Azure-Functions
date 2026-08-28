declare interface IRecentDocumentsWebPartStrings {
  Settingproperties: string;
  Title: string;
  DocsToShow: string;
  DocsToShowRequired: string;
  MoreLink: string;
  MoreLinkLabel: string;
  NoDocumentsAvailable: string;
  ViewMoreRelevantDocuments: string;
}

declare module 'DocumentsWebPartStrings' {
  const strings: IDocumentsWebPartStrings;
  export = strings;
}