import { Buttonprops } from "@fluentui/react-components";

export interface IConfirmActionprops {
    dialogTitle: any;
    dialogContent: any;
    dialogTrigger?: any;
    loadingActionContent?: any;
    dialogprops?: { open: boolean };
    cancelButtonprops: Buttonprops;
    acceptButtonprops: Buttonprops;
}

