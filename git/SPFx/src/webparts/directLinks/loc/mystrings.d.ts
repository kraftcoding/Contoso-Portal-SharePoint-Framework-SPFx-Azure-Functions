declare interface IDirectLinksWebPartStrings {
  Settingproperties: string;
  Title: string;
  SlidesToShow: string;
  SlidesToScroll: string;
  SlidesToShowRequired: string;
  SlidesToScrollRequired: string;
  ScrollLessThanOrEqualToShow: string;
  LoadingTheUsefulLinks: string;
  NoUsefulLinksAvailable: string;
  SeeMore: string;
  SlideInformation: string;
  MoveRight: string;
  MoveLeft: string;
}

declare module 'DirectLinksWebPartStrings' {
  const strings: IDirectLinksWebPartStrings;
  export = strings;
}