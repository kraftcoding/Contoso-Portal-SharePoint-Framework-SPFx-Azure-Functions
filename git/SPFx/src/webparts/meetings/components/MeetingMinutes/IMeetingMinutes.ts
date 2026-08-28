import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEvent, IUserAttendance } from "../../../../service/BackendServiceModels/EventModels";
import SPService from "../../../../service/Service";
import { BackendService } from "../../../../service/BackendService";
import { IFileInfo } from "@pnp/sp/files";

export interface IMeetingMinutesprops {
    event: IEvent;
    isEditor: boolean;
    context: WebPartContext;
    spService: SPService;
    bkService: BackendService;
    readOnly: boolean;
}

export interface IMeetingMinutesState {
    isLoading: boolean;
    minutesSections: IMinutesSection[];
    selectedMinuteSection?: IMinutesSection;
    selectedMinuteSectionFormErrorMessage: string;
    buttonToSaveTheSelectedMinuteSectionDisabled: boolean;
    openAttendanceDialog: boolean;
    openDocDialog: boolean;
    approdvalStatus?: IApprodvalStatusOfTheMinutes;
    attendances: IUserAttendance[];
    minutesFile?: IFileInfo;
    dialogLoading: boolean;
    downloadingTemplate: boolean;
    showApprodvalDetailModal: boolean;
}

export interface IMinutesSection {
    Id: string;
    Title: string;
    Icon: React.ReactElement;
    Description: string;
    Url: string;
}

export enum MinutesSectionsIds {
    AgendaInformation = 'AgendaInformation',
    AgreementsInformation = 'AgreementsInformation',
    AttendanceInformation = 'AttendanceInformation',
    DocumentationInformation = 'DocumentationInformation'
}

export interface IBackMinutes {
    IdMinutes?: string;
    MinutesStatus?: string;
    AgendaInformation: string;
    AgreementsInformation: string;
    AttendanceInformation: string;
    DocumentationInformation: string;
    Attendances: IUserAttendance[];
}

export interface IApprodvalStatusOfTheMinutes {
    ApprodvalsPendings: IApprodvalStatusUserMinutes[];
    ApprodvalsAccepted: IApprodvalStatusUserMinutes[];
    ApprodvalsRejected: IApprodvalStatusUserMinutes[];
    ApprodvalsPendingModification: IApprodvalStatusUserMinutes[];
    ApprodvalEndDate: string;
}

export interface IApprodvalStatusUserMinutes{
    Upn: string;
    FullName: string;
}
