declare interface IAgendaAnnouncementWebPartStrings {
  Title: string;
  Today: string;
  See: string;
  AddOutlook: string;
  Link: string;
  GoTotestviousMonth: string;
  GoToToday: string;
  GoToNextMonth: string;
  ThereAreMeetingsOnDay: string;
}

declare module 'AgendaAnnouncementWebPartStrings' {
  const strings: IAgendaAnnouncementWebPartStrings;
  export = strings;
}