declare interface ILaunchDataprocessCommandSetStrings {
  ModalTitle: string;
  CompleteLabel: string;
  IncrementalLabel: string;
  processType: string;
  AreYouSure: string;
  RunningMsg: string;
  ErrorMsg: string;
  LaunchedMsg: string;
  CancelButton: string;
  AcceptButton: string;
}

declare module 'LaunchDataprocessCommandSetStrings' {
  const strings: ILaunchDataprocessCommandSetStrings;
  export = strings;
}
