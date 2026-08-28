import SPService from "../../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../../service/BackendService";
import { IEvent, IUserAttendance } from "../../../../service/BackendServiceModels/EventModels";

export interface IMeetingAssistant {
    Id: string;
    CompletedName: string;
    State: string;
    IsChecked: boolean;
    CanVote: boolean;
}

export interface IMeetingAssistantsprops {
    context: WebPartContext;
    spService: SPService;
    bkService: IBackendService;
    isEditor: boolean;
    readOnly: boolean;
    event: IEvent;
    match: {
        params: {
            id: any;
        }
    };
}

export interface IMeetingAssistantsState {
    attendances: IUserAttendance[];
    isLoading: boolean;
    form?: {
        UserPrincipalName?: string;
        isGuest?: boolean;
    };
    isSaving: boolean;
    openAdd: boolean;
    userError: string;
    editing: boolean;
    loadingAttendance?: string;
    downloadingAttendance: boolean;
    groupId: string;
    reportingAttendance: boolean;
    reportingAttendanceErrorMessage: string;
    membersGroupId: string;
    guestsGroupId: string;
    convoGroupId: string;
    gestorGroupId: string;
    asistenteGroupId: string;
}