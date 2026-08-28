declare interface ISendDocToBodyCommandSetStrings {
  Command1: string;
  DialogTitle: string;
  DropDownPlaceholder: string;
  ButtonCancelar: string;
  ButtonAceptar: string;
  MessageOK: string;
  MessageCargando: string;
  CommandText: string;
  ErrorTextOnLoad: string;
  ErrorTextOnSave: string;
}

declare module 'SendDocToBodyCommandSetStrings' {
  const strings: ISendDocToBodyCommandSetStrings;
  export = strings;
}
