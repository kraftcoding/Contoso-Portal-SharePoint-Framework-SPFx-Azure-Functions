declare interface IPendingTasksWebPartStrings {
  NumberOfResults: string;
  GetPendingTasksFromThisSite: string;
  Title: string;
  MoreLink: string;
  MoreLinkLabel: string;
  LoadingThePendingTasks: string;
  NoPendingTasksToShow: string;
  MeetingAttendance: string;
  AttendanceType: string;
  DelegateTheMeeting: string;
  MinutesApprodval: string;
  MinutesModification: string;
  CertificationApprodval: string;
  Attendance: string;
  Vote: string;
  SelectAnOption: string;
  SelectAPerson: string;
  AdditionalComments: string;
  WriteAComment: string;
  RuleOut: string;
  Send: string;
  Sending: string;
  Accept: string;
  Refuse: string;
  Delegate: string;
  Online: string;
  InPerson: string;
  ModifyMinutes: string;
  DropdownErrorMessage: string;
  PeoplePickerErrorMessage: string;
  PeoplePickerDelegationErrorMessage: string;
  CommentsErrorMessage: string;
  BackendErrorMessage: string;
  prefixMeetingAttendanceRequest: string;
  prefixDelegatedMeetingAttendanceRequest: string;
  prefixMinutesApprodvalRequest: string;
  prefixMinutesModificationRequest: string;
  prefixCertificationRequest: string;
  DelegateAssistanceTo: string;
  DelegateAssistanceAndVoteTo: string;
  DelegateVoteTo: string;
  RequestToModifyTheMinutes: string;
  Reason: string;
  ThereIsACertificationRequest: string;
  ViewMorePendingTasks: string;
}

declare module 'PendingTasksWebPartStrings' {
  const strings: IPendingTasksWebPartStrings;
  export = strings;
}