declare interface ImyDepartmentsWebPartStrings {
  Settingproperties: string;
  Title: string;
  SlidesToShow: string;
  SlidesToScroll: string;
  SlidesToShowRequired: string;
  SlidesToScrollRequired: string;
  ScrollLessThanOrEqualToShow: string;
  LoadingmyDepartments: string;
  NomyDepartmentsAvailable: string;
  SlideInformation: string;
  MoveRight: string;
  MoveLeft: string;
}

declare module 'myDepartmentsWebPartStrings' {
  const strings: ImyDepartmentsWebPartStrings;
  export = strings;
}
