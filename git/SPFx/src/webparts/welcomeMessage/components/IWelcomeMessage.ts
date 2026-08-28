import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ISPService } from "../../../service/Service";
import { IBackendService } from "../../../service/BackendService";
import { Term } from '../../../models/ITag';
import { INotification } from "../../../service/BackendServiceModels/NotificationModel";
import { Iprofile } from "../../../service/BackendServiceModels/profileModel";
import { IPendingChanges } from "../../../models/IPendingChanges";

export interface IWelcomeMessageprops {
    context: WebPartContext;
    spService: ISPService;
    bkService: IBackendService;
    locale: string;
    getNotificationsFromAllSites: boolean;
}

export interface IWelcomeMessageState {
    loadingUserInformation: boolean;
    savingUserInformation: boolean;
    userInformation?: Iprofile;
    currentAlerts?: INotification[];
    archivedAlerts?: INotification[];
    loadingArchivedAlerts: boolean;
    autonomousCommunities: Term[];
    attendanceTypes: Term[];
    onlineTools: Term[];
    dialogIsOpen: boolean;
    popoverIsOpen: boolean;
    drawerIsOpen: boolean;
    backendErrorMessage: string;
    backendHasFailed: boolean;
    pendingChanges?: IPendingChanges[];
    selectedCurrentAlert?: INotification;
    selectedArchivedAlert?: INotification;
    archivingMeetings: boolean;
}