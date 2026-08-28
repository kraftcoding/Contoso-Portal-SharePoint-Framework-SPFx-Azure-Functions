import { ISPService } from "../../../service/Service";
import { IBackendService } from "../../../service/BackendService";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserTaskModel } from "../../../service/BackendServiceModels/UserTaskModel";

export interface IPendingTasksprops {
  context: WebPartContext;
  spService: ISPService;
  bkService: IBackendService;
  locale: string;
  title: string;
  moreLink: string;
  numberOfResults: number;
  getThePendingTasksFromThisSite: boolean;
}

export interface IPendingTasksState {
  loadingThePendingTasks: boolean;
  pendingTasks: IUserTaskModel[];
}