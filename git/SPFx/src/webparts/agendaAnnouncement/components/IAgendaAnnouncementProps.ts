import { IBackendService } from "../../../service/BackendService";
import { ISPService } from "../../../service/Service";

export interface IAgendaAnnouncementprops {
  title: string;
  bkService: IBackendService;
  spService: ISPService;
  locale: string;
}
