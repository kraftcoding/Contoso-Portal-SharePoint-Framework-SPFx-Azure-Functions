import { DisplayMode } from "@microsoft/sp-core-library";
import { ISPService } from "../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEvent } from "../../../service/BackendServiceModels/EventModels";
import { IBackendService } from "../../../service/BackendService";
import { Term } from "../../../models/ITag";

export interface IMeetingsprops {
  spService: ISPService;
  bkService: IBackendService;
  title: string;
  showVoting: boolean;
  displayMode: DisplayMode;
  updateproperty: (value: string) => void;
  context: WebPartContext;
  locale: string;
}

export interface IMeetingsState {
  events: IEvent[];
  isLoading: boolean;
  selectedEvent?: IEvent;
  statusTerms: Term[];
  loadingChangedEvent: boolean;
  loadingChangedEventId?: string;
  isEditor: boolean;
  isGuest: boolean;
  selectedEventReadOnly: boolean;
  showArchived: boolean;
  isOCprodle: boolean;
}