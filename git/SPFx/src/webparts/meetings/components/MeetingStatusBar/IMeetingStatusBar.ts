import { ISPService } from "../../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../../service/BackendService";
import { EventStatus, IEvent } from "../../../../service/BackendServiceModels/EventModels";
import { Term } from "../../../../models/ITag";
import { IPendingChanges } from "../../../../models/IPendingChanges";

export interface IMeetingStatusBarprops {
  context: WebPartContext;
  spService: ISPService;
  bkService: IBackendService;
  event?: IEvent;
  locale: string;
  statusTerms: Term[];
  onStatusChange: (nextStatusId: string) => void;
  onDataChanged: (changedEvent: IEvent) => void;
  loadingStatus: boolean;
  loadingEventId?: string;
  isEditor: boolean;
  readOnly: boolean;
  isOCprodle: boolean;
}

export interface IMeetingStatusBarState {
  openEditForm: boolean;
  updatesAvailable: boolean;

  sendUpdateDialogIsOpen: boolean;
  sendUpdateErrMsg: string;

  urgentNotifDialogOpen: boolean;
  urgentNotifErrMsg: string;
  textToSend: string;
  sendingUrgentNotif: boolean;

  dialogOpen: boolean,
  ErrMsg:string,
  nextState? : EventStatus;

  pendingChanges?: IPendingChanges[];

  attendanceTypes: Term[];
  onlineTools: Term[];
}