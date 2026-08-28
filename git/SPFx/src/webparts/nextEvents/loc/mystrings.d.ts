declare interface INextEventsWebPartStrings {
  Settingproperties: string;
  Title: string;
  SlidesToShow: string;
  SlidesToScroll: string;
  SlidesToShowRequired: string;
  SlidesToScrollRequired: string;
  ScrollLessThanOrEqualToShow: string;
  LoadingTheNextEvents: string;
  NoNextEventsAvailable: string;
  AddEvent: string;
  MoreLink: string;
  MoreLinkLabel: string;
  SlideInformation: string;
  GoToEvent: string;
  GoToTheEventDetail: string;
  ViewMoreUpcomingEvents: string;
  MoveRight: string;
  MoveLeft: string;
}

declare module 'NextEventsWebPartStrings' {
  const strings: INextEventsWebPartStrings;
  export = strings;
}