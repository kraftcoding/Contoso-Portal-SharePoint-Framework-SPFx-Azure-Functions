/* eslint-disable @typescript-eslint/no-explicit-any */
import { Term } from '../../../../models/ITag';
import { ISPService } from '../../../../service/Service';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IBackendService } from '../../../../service/BackendService';
import { IUserTaskModel } from '../../../../service/BackendServiceModels/UserTaskModel';

/* Request de Attendance a Meeting */

export interface IMeetingAttendanceRequestprops {
    locale: string;
    context: WebPartContext;
    spService: ISPService;
    bkService: IBackendService;
    organTaxonomyTerms: Term[];
    pendingTask: IUserTaskModel;
    getThePendingTasks(): promise<void>;
    getThePendingTasksFromThisSite: boolean;
    DepartmentName: string;
}

export interface IMeetingAttendanceRequestState {
    formInformation?: IMeetingAttendanceRequest;
    meetingAttendanceDropdownErrorMessage: string;
    attendanceTypeDropdownErrorMessage: string;
    delegatedAttendancePeoplePickerErrorMessage: string;
    delegatedVotePeoplePickerErrorMessage: string;
    commentsErrorMessage: string;
    backendErrorMessage: string;
    sendingTheFormInformation: boolean;
    sendButtonIsDisabled: boolean;
    membersGroupId: string;
    asistenteGroupId: string;
    currentTask:IUserTaskModel;
    isLoadingDetail:boolean;
}

export interface IMeetingAttendanceRequest {
    selectedOptionFromTheMeetingAttendanceDropdown?: string;
    selectedTermIdFromTheMeetingAttendanceDropdown?: string;
    selectedOptionFromTheAttendanceTypeDropdown?: string;
    selectedTermIdFromTheAttendanceTypeDropdown?: string;
    attendanceIsDelegated?: boolean;
    voteIsDelegated?: boolean;
    personSelectedToDelegateTheAttendance?: any[];
    personSelectedToDelegateTheVote?: any[];
    additionalComments?: string;
}

/* Request de delegación de Attendance a Meeting */

export interface IDelegatedMeetingAttendanceRequestprops {
    locale: string;
    bkService: IBackendService;
    organTaxonomyTerms: Term[];
    pendingTask: IUserTaskModel;
    getThePendingTasks(): promise<void>;
    getThePendingTasksFromThisSite: boolean;
    DepartmentName: string;
}

export interface IDelegatedMeetingAttendanceRequestState {
    formInformation?: IDelegatedMeetingAttendanceRequest;
    meetingAttendanceDropdownErrorMessage: string;
    commentsErrorMessage: string;
    backendErrorMessage: string;
    sendingTheFormInformation: boolean;
    sendButtonIsDisabled: boolean;
    currentTask:IUserTaskModel;
    isLoadingDetail:boolean;
}

export interface IDelegatedMeetingAttendanceRequest {
    selectedOptionFromTheMeetingAttendanceDropdown?: string;
    selectedTermIdFromTheMeetingAttendanceDropdown?: string;
    additionalComments?: string;
}

/* Request de aprodbación de Minutes */

export interface IMinutesApprodvalRequestprops {
    locale: string;
    bkService: IBackendService;
    organTaxonomyTerms: Term[];
    pendingTask: IUserTaskModel;
    getThePendingTasks(): promise<void>;
    getThePendingTasksFromThisSite: boolean;
    DepartmentName: string;
}

export interface IMinutesApprodvalRequestState {
    formInformation?: IMinutesApprodvalRequest;
    minutesApprodvalDropdownErrorMessage: string;
    commentsErrorMessage: string;
    backendErrorMessage: string;
    sendingTheFormInformation: boolean;
    sendButtonIsDisabled: boolean;
    currentTask:IUserTaskModel;
    isLoadingDetail:boolean;
}

export interface IMinutesApprodvalRequest {
    selectedOptionFromTheMinutesApprodvalDropdown?: string;
    selectedTermIdFromTheMinutesApprodvalDropdown?: string;
    additionalComments?: string;
}

/* Request de modificación de Minutes */

export interface IMinutesModificationRequestprops {
    locale: string;
    bkService: IBackendService;
    organTaxonomyTerms: Term[];
    pendingTask: IUserTaskModel;
    getThePendingTasks(): promise<void>;
    getThePendingTasksFromThisSite: boolean;
    DepartmentName: string;
}

export interface IMinutesModificationRequestState {
    formInformation?: IMinutesModificationRequest;
    minutesModificationDropdownErrorMessage: string;
    commentsErrorMessage: string;
    backendErrorMessage: string;
    sendingTheFormInformation: boolean;
    sendButtonIsDisabled: boolean;
    currentTask:IUserTaskModel;
    isLoadingDetail:boolean;
}

export interface IMinutesModificationRequest {
    selectedOptionFromTheMinutesModificationDropdown?: string;
    selectedTermIdFromTheMinutesModificationDropdown?: string;
    additionalComments?: string;
}

/* Request de certificación */

export interface ICertificationRequestprops {
    locale: string;
    bkService: IBackendService;
    organTaxonomyTerms: Term[];
    pendingTask: IUserTaskModel;
    getThePendingTasks(): promise<void>;
    getThePendingTasksFromThisSite: boolean;
    DepartmentName: string;
}

export interface ICertificationRequestState {
    formInformation?: ICertificationRequest;
    certificationApprodvalDropdownErrorMessage: string;
    backendErrorMessage: string;
    sendingTheFormInformation: boolean;
    sendButtonIsDisabled: boolean;
    currentTask:IUserTaskModel;
    isLoadingDetail:boolean;
}

export interface ICertificationRequest {
    selectedOptionFromTheCertificationApprodvalDropdown?: string;
    selectedTermIdFromTheCertificationApprodvalDropdown?: string;
}