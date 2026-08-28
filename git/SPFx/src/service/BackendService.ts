import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { AadHttpClient, AadHttpClientFactory } from '@microsoft/sp-http';
import { EnvConfig } from './../utils/EnvConfig';
import { PageContext } from '@microsoft/sp-page-context';
import { Logger } from '../utils/Logger';
import { BackendErrorResponse } from './BackendServiceModels/IBackendErrorResponse';
import { IEvent, IEventAgendaItem, IEventAgreement, IEventSendOperationInformation, IEventStorage, IEventVotationDetails, INewOrUpdatedEvent, IUserAttendance, IVote, ITenantSensitivityLabel, IDocumentSensitivityLabel, IEventFromBbdd } from './BackendServiceModels/EventModels';
import { IBody } from './BackendServiceModels/BodyModel';
import { IMember } from './BackendServiceModels/MemberModel';
import { Iprofile } from './BackendServiceModels/profileModel';
import { ITaskCertification, IUserTaskDelegateModel, IUserTaskModel } from './BackendServiceModels/UserTaskModel';
import { IUserTaskResponse } from './BackendServiceModels/UserTaskResponse';
import { INotification, INotificationBathRequest } from './BackendServiceModels/NotificationModel';
import { IApprodvalStatusOfTheMinutes, IBackMinutes } from '../webparts/meetings/components/MeetingMinutes/IMeetingMinutes';
import { IPendingChanges } from '../models/IPendingChanges';
import {BodyUserInformation, ILicenseInformation, IdepartmentManagement, IOrganEXTERNAL, IRelationBusinessArea, IUserManagement, IUsersOrganRelation } from './BackendServiceModels/ManagementModel';
import { IPermissionCheckResult } from '../webparts/administrationApp/models/AdministrationAppModels';

export interface IBackendService {
    getUserBodies(): promise<IBody[]>;
    getUserEvents(fromArchive?: boolean): promise<IEvent[]>;
    ping(): promise<string>;
    createEvent(bodyId: string, event: INewOrUpdatedEvent): promise<IEvent>;
    getEvents(bodyId?: string, fromArchive?: boolean): promise<IEvent[]>;
    getEvent(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEvent>;
    updateEvent(bodyId: string, eventId: string, event: INewOrUpdatedEvent): promise<IEvent>;
    getEventStorage(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventStorage>;
    sendEventById(bodyId: string, eventId: string, message: string): promise<IEventSendOperationInformation>;
    eventPendingChanges(bodyId: string, eventId: string): promise<IPendingChanges[]>;
    changeEventStatusById(bodyId: string, eventId: string, newEventStatusId: string, message: string): promise<IEventSendOperationInformation>;
    getEventAgendaItems(bodyId: string, id: string, fromArchive?: boolean): promise<IEventAgendaItem[]>;
    addOrUpdateEventAgendaItems(bodyId: string, eventId: string, agendaItems: IEventAgendaItem[]): promise<IEventAgendaItem[]>;
    deleteEventAgendaItem(bodyId: string, eventId: string, agendaItemId: string): promise<void>;
    addRelatedDocumentByUniqueId(bodyId: string, eventId: string, agendaItemId: string, docUniqueId: string): promise<void>;
    removeRelatedDocumentByUniqueId(bodyId: string, eventId: string, agendaItemId: string, docUniqueId: string): promise<void>;
    addRelatedDocumentIdToAgreement(bodyId: string, eventId: string, agreementItemId: string, docUniqueId: string, isCertificate: boolean): promise<void>;
    removeRelatedDocumentIdToAgreement(bodyId: string, eventId: string, agreementItemId: string, docUniqueId: string): promise<void>;
    deleteDocumentByUniqueId(bodyId: string, eventId: string, docUniqueId: string): promise<void>;
    getUserTask(): promise<IUserTaskModel[]>;
    getUserTaskByBodyId(bodyId: string): promise<IUserTaskModel[]>;
    getBodyUsersByRole(bodyId: string, role: string): promise<IMember[]>;
    updateTask(bodyId: string, taskId: string, termId: string, attendanceFormatId: string, comment: string): promise<IUserTaskResponse[]>;
    getUserNotifications(source: string, locale: string): promise<INotification[]>;
    getArchivedNotifications(source: string, locale: string): promise<INotification[]>;
    getNotificationInfo(notificationId: string): promise<INotification>;
    setReadNotification(notificationId: string, readed: boolean): promise<void>;
    hideNotification(notificationId: string): promise<void>;
    updateNotificationsInBatch(request: INotificationBathRequest): promise<void>;
    addUrgentNotification(notification: INotification, bodyId: string, eventId: string): promise<void>;
    getCurrentUserprofile(): promise<Iprofile>;
    updateCurrentUserprofile(userInfo: Iprofile): promise<void>;
    getAttendanceByEventId(bodyId: string, eventId: string, fromArchive?: boolean): promise<IUserAttendance[]>;
    addOrUpdateAttendance(bodyId: string, eventId: string, attendance: string): promise<IUserAttendance[]>;
    deleteAttendance(bodyId: string, eventId: string, attendance: string): promise<IUserAttendance[]>;
    updateAttendance(bodyId: string, eventId: string, attendances: IUserAttendance[]): promise<void>;
    createDelegateTask(bodyId: string, taskDelegate: IUserTaskDelegateModel): promise<void>;
    updateTaskDelegate(bodyId: string, taskId: string, termId: string, comment: string): promise<void>;
    getAgreementsByEventId(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventAgreement[]>;
    getPendingCertificateAgreementsByEventId(bodyId: string, eventId: string): promise<string[]>;
    addOrUpdateEventAgreements(bodyId: string, eventId: string, agreement: IEventAgreement[]): promise<IEventAgreement[]>;
    updateMinutesInformation(bodyId: string, eventId: string, updatedMinutes: IBackMinutes): promise<void>;
    approdvalTaskMinutes(bodyId: string, taskId: string, termId: string, comment: string): promise<void>;
    modificationTaskMinutes(bodyId: string, taskId: string, termId: string, comment: string): promise<void>;
    startTheMinuteApprodvalprocess(bodyId: string, eventId: string): promise<void>;
    requestCertificate(bodyId: string, taskCertification: ITaskCertification): promise<void>;
    approdvalTaskCertification(bodyId: string, taskId: string, termId: string): promise<void>;
    getApprodvalStatusMinutes(bodyId: string, eventId: string): promise<IApprodvalStatusOfTheMinutes>;
    addRelatedDocumentIdToMinutes(bodyId: string, eventId: string, docUniqueId: string): promise<void>;
    getMinutesInfo(bodyId: string, eventId: string, fromArchive?: boolean): promise<IBackMinutes>;
    downloadAgreementsTemplate(bodyId: string, eventId: string, agreementId: string): promise<ArrayBuffer>;
    downloadOrderOfTheDayTemplate(bodyId: string, eventId: string): promise<ArrayBuffer>;
    downloadMinutesTemplate(bodyId: string, eventId: string): promise<ArrayBuffer>;
    downloadAttendanceTemplate(bodyId: string, eventId: string): promise<ArrayBuffer>;
    getEventVotationDetail(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventVotationDetails[]>;
    addOrUpdateEventVotationDetail(bodyId: string, eventId: string, votations: IEventVotationDetails[]): promise<IEventVotationDetails[]>;
    getEventVotingRetestsentation(bodyId: string, eventId: string, fromArchive?: boolean): promise<Record<string, string[]>>;
    startEventVotation(bodyId: string, eventId: string): promise<IEventVotationDetails[]>;
    updateVoteByUser(bodyId: string, eventId: string, vote: IVote): promise<IEventVotationDetails[]>;
    updateVoteByRetestsentation(bodyId: string, eventId: string, retestsentationId: string, vote: IVote): promise<IEventVotationDetails>;
    getOrganInfo(bodyId: string): promise<IdepartmentManagement[]>;
    getUsersInfo(bodyId: string): promise<Record<string, IUserManagement>>;
    getUserTaskByTaskId(taskId:string, bodyId: string): promise<IUserTaskModel>;
    updateUserprofile(userInfo: IUserManagement, bodyId: string): promise<void>;
    updateOrganInfo(organInfo: IdepartmentManagement, bodyId: string): promise<void>;
    updateUsersInOrgan(membersInfo: IUsersOrganRelation[], bodyId: string): promise<void>;
    checkUserHasPermissions(checkInfo: IUsersOrganRelation): promise<IPermissionCheckResult>;
    getEXTERNALOrgansInfo(bodyId: string): promise<IOrganEXTERNAL[]>;
    updateEXTERNALOrganInfo(organInfo: IOrganEXTERNAL, bodyId: string): promise<void>;
    createEXTERNALOrgan(organInfo: IOrganEXTERNAL, bodyId: string): promise<void>;
    downloadBodyCertificate(bodyId: string, body:string, isEXTERNAL?:boolean): promise<ArrayBuffer>;
    getAreaRelationsInfo(bodyId: string): promise<IRelationBusinessArea[]>;
    updateAreaRelationsInfo(item:IRelationBusinessArea, bodyId: string):promise<void>;
    getAvailableLicenses(bodyId:string):promise<ILicenseInformation[]>;
    manageUserLicense(bodyId:string, upn:string, license:string):promise<void>;
    createUser(item:BodyUserInformation, bodyId: string):promise<void>;
    getSensitivityLabels():promise<ITenantSensitivityLabel[]>;
    getDocumentsSensitivityLabelsByEventSharedId(bodyId:string, eventId: string, fromArchive?:boolean, isEditor?:boolean):promise<IDocumentSensitivityLabel[]>;
    updateSensitivityLabel(documentSensitivityLabel:IDocumentSensitivityLabel, fromArchive?:boolean, isEditor?:boolean):promise<boolean>;
    launchSaveDataprocess(startDate?:string): promise<string>;
    getUserInvitationMessage(serverRelativeUrl:string): promise<string>;
    getAllEventsFromBbdd(startDate:string, endDate:string): promise<IEventFromBbdd[]>;
}

export class BackendService implements IBackendService {
    public static readonly serviceKey: ServiceKey<IBackendService> = ServiceKey.create<IBackendService>(
        'SPFx:ContosoBackendService',
        BackendService
    );

    private _aadHttpClientFactory: AadHttpClientFactory;
    private _aadHttpClient: promise<AadHttpClient>;
    private _pageContext: PageContext;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._aadHttpClientFactory = serviceScope.consume(AadHttpClientFactory.serviceKey);
            this._aadHttpClient = this._aadHttpClientFactory.getClient(EnvConfig.BakendAppClientId);
            this._pageContext = serviceScope.consume(PageContext.serviceKey);
        });
    }

    public async ping(): promise<string> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/ping', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to ping', error, this._pageContext);
            throw error;
        }
    }

    public async getEvents(bodyId?: string, fromArchive?: boolean): promise<IEvent[]> {
        const client = await this._aadHttpClient;
        try {
            const apiUrl = (bodyId) ? '/api/bodies/' + bodyId + '/events' : '/api/events';
            const response = await client.get(EnvConfig.BackendHost + apiUrl + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get events', error, this._pageContext);
            throw error;
        }
    }

    public async getEvent(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEvent> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get event:' + eventId, error, this._pageContext);
            throw error;
        }
    }

    public async getUserEvents(fromArchive?: boolean): promise<IEvent[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/events' + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get events:', error, this._pageContext);
            throw error;
        }
    }

    public async createEvent(bodyId: string, event: INewOrUpdatedEvent): promise<IEvent> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events', AadHttpClient.configurations.v1, { body: JSON.stringify(event) });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to create event', error, this._pageContext);
            throw error;
        }
    }

    public async updateEvent(bodyId: string, eventId: string, event: INewOrUpdatedEvent): promise<IEvent> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId, AadHttpClient.configurations.v1, { body: JSON.stringify(event), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to update event', error, this._pageContext);
            throw error;
        }
    }

    public async sendEventById(bodyId: string, eventId: string, message: string): promise<IEventSendOperationInformation> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/send?changeMessage=" + encodeURIComponent(message), AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to send', error, this._pageContext);
            throw error;
        }
    }

    public async eventPendingChanges(bodyId: string, eventId: string): promise<IPendingChanges[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/pendingChanges", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            // const responseJson = await response.json();
            // return responseJson.result;
            return await response.json();
        } catch (error) {
            Logger.error('Failed to verify pending changes', error, this._pageContext);
            throw error;
        }
    }

    public async changeEventStatusById(bodyId: string, eventId: string, newEventStatusId: string, message: string): promise<IEventSendOperationInformation> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/changeStatus?statusId=" + newEventStatusId + "&changeMessage=" + encodeURIComponent(message), AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to change status', error, this._pageContext);
            throw error;
        }
    }

    public async getEventStorage(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventStorage> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/docs" + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get event storage:' + eventId, error, this._pageContext);
            throw error;
        }
    }

    public async getEventAgendaItems(bodyId: string, id: string, fromArchive?: boolean): promise<IEventAgendaItem[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + id + "/items" + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get event agenda items:' + id, error, this._pageContext);
            throw error;
        }
    }

    public async addOrUpdateEventAgendaItems(bodyId: string, eventId: string, agendaItems: IEventAgendaItem[]): promise<IEventAgendaItem[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items", AadHttpClient.configurations.v1, { body: JSON.stringify(agendaItems), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to add or update agenda items', error, this._pageContext);
            throw error;
        }
    }

    public async deleteEventAgendaItem(bodyId: string, eventId: string, agendaItemId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items/" + agendaItemId, AadHttpClient.configurations.v1, { method: 'DELETE' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to delete agenda item', error, this._pageContext);
            throw error;
        }
    }

    public async getUserBodies(): promise<IBody[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/my', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get bodies user', error, this._pageContext);
            throw error;
        }
    }

    public async getBodyUsersByRole(bodyId: string, role: string): promise<IMember[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/users?role=' + role, AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get body user', error, this._pageContext);
            throw error;
        }
    }

    public async addRelatedDocumentByUniqueId(bodyId: string, eventId: string, agendaItemId: string, docUniqueId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items/" + agendaItemId + '/addRelatedDocumentByUniqueId?docUniqueId=' + docUniqueId, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to add related document', error, this._pageContext);
            throw error;
        }
    }

    public async removeRelatedDocumentByUniqueId(bodyId: string, eventId: string, agendaItemId: string, docUniqueId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items/" + agendaItemId + '/removeRelatedDocumentByUniqueId?docUniqueId=' + docUniqueId, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to remove related document', error, this._pageContext);
            throw error;
        }
    }

    public async addRelatedDocumentIdToAgreement(bodyId: string, eventId: string, agreementItemId: string, docUniqueId: string, isCertificate: boolean): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items/" + agreementItemId + '/addRelatedDocumentIdToAgreement?docUniqueId=' + docUniqueId + "&isCertificate=" + isCertificate, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to add related document', error, this._pageContext);
            throw error;
        }
    }

    public async removeRelatedDocumentIdToAgreement(bodyId: string, eventId: string, agreementItemId: string, docUniqueId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/items/" + agreementItemId + '/removeRelatedDocumentIdToAgreement?docUniqueId=' + docUniqueId, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to remove related document', error, this._pageContext);
            throw error;
        }
    }

    public async addRelatedDocumentIdToMinutes(bodyId: string, eventId: string, docUniqueId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/SetMinutesDocument?docUniqueId=' + docUniqueId, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to add related document', error, this._pageContext);
            throw error;
        }
    }

    public async deleteDocumentByUniqueId(bodyId: string, eventId: string, docUniqueId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/deleteDocumentByUniqueId?docUniqueId=' + docUniqueId, AadHttpClient.configurations.v1, { body: "" });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to delete related document', error, this._pageContext);
            throw error;
        }
    }

    public async getMinutesInfo(bodyId: string, eventId: string, fromArchive?: boolean): promise<IBackMinutes> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/minutesinformation' + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get bodies user', error, this._pageContext);
            throw error;
        }
    }

    public async getUserTask(): promise<IUserTaskModel[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/tasks', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get user task:', error, this._pageContext);
            throw error;
        }
    }

    public async getUserTaskByBodyId(bodyId: string): promise<IUserTaskModel[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId, AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get user task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getUserNotifications(source: string, locale: string): promise<INotification[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/notifications?source=' + source + '&locale=' + locale, AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to get user notifications from source ${source} and locale ${locale}`, error, this._pageContext);
            throw error;
        }
    }

    public async getArchivedNotifications(source: string, locale: string): promise<INotification[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/notificationsArchived?source=' + source + '&locale=' + locale, AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to get archived notifications from source ${source} and locale ${locale}`, error, this._pageContext);
            throw error;
        }
    }

    public async setReadNotification(notificationId: string, setStatus: boolean): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/notifications/' + notificationId + '/setStatus/' + setStatus, AadHttpClient.configurations.v1, { method: 'PATCH' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to set user notifications from id: ${notificationId}`, error, this._pageContext);
            throw error;
        }
    }

    public async hideNotification(notificationId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/notifications/' + notificationId + '/hide', AadHttpClient.configurations.v1, { method: 'PATCH' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to hide user notifications from id: ${notificationId}`, error, this._pageContext);
            throw error;
        }
    }

    public async updateNotificationsInBatch(request: INotificationBathRequest): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/notifications/batch', AadHttpClient.configurations.v1, { body: JSON.stringify(request), method: 'PATCH' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to batch user notifications from request: ${request}`, error, this._pageContext);
            throw error;
        }
    }

    public async addUrgentNotification(notification: INotification, bodyId: string, eventId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/notifications/bodies/' + bodyId + '/event/' + eventId + '/UrgentNotification', AadHttpClient.configurations.v1, { body: JSON.stringify(notification), method: 'POST' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to add event notification from request: ${eventId}`, error, this._pageContext);
            throw error;
        }
    }

    public async getNotificationInfo(notificationId: string): promise<INotification> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/notifications/' + notificationId, AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to get user notifications by id: ${notificationId}`, error, this._pageContext);
            throw error;
        }
    }

    public async getCurrentUserprofile(): promise<Iprofile> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/profile', AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async updateCurrentUserprofile(userInfo: Iprofile): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/profile', AadHttpClient.configurations.v1, { body: JSON.stringify(userInfo), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async getAttendanceByEventId(bodyId: string, eventId: string, fromArchive?: boolean): promise<IUserAttendance[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/attendance' + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async addOrUpdateAttendance(bodyId: string, eventId: string, attendance: string): promise<IUserAttendance[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/addAttendance/' + attendance, AadHttpClient.configurations.v1, { method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to add attendance in list`, error, this._pageContext);
            throw error;
        }
    }

    public async deleteAttendance(bodyId: string, eventId: string, attendance: string): promise<IUserAttendance[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/deleteAttendance/' + attendance, AadHttpClient.configurations.v1, { method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error(`Failed to delete attendance in list`, error, this._pageContext);
            throw error;
        }
    }

    public async updateAttendance(bodyId: string, eventId: string, attendances: IUserAttendance[]): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/updateAttendance', AadHttpClient.configurations.v1, { body: JSON.stringify(attendances), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to update attendance in list`, error, this._pageContext);
            throw error;
        }
    }

    public async getAgreementsByEventId(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventAgreement[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/agreements" + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get agreement items:' + eventId, error, this._pageContext);
            throw error;
        }
    }

    // "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/pendingcertagreements")] HtttestquestData req, string bodyId, string eventId)

    public async getPendingCertificateAgreementsByEventId(bodyId: string, eventId: string): promise<string[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/pendingcertagreements", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get pendingCertificate agreement items:' + eventId, error, this._pageContext);
            throw error;
        }
    }

    public async addOrUpdateEventAgreements(bodyId: string, eventId: string, agreements: IEventAgreement[]): promise<IEventAgreement[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/agreements", AadHttpClient.configurations.v1, { body: JSON.stringify(agreements), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to add or update agreement items', error, this._pageContext);
            throw error;
        }
    }

    private static async handleServerSideException(
        response: any, // eslint-disable-line @typescript-eslint/no-explicit-any
        throwOnError: boolean = true,
    ): promise<void | BackendErrorResponse> {
        if (!response.ok) {
            let jsonResponse: BackendErrorResponse = {} as BackendErrorResponse;
            try {
                jsonResponse = await response.json();
            } catch {
                // ignore
            }
            if (throwOnError) {
                if (jsonResponse.exceptionType && jsonResponse.exceptionType.trim())
                    throw new Error("Backend exception '" + jsonResponse.exceptionType + "': " + jsonResponse.message);
                else throw new Error("Backend service error '" + response.status + "': " + response.statusText);
            } else {
                if (jsonResponse.exceptionType && jsonResponse.exceptionType.trim())
                    return jsonResponse as BackendErrorResponse;
            }
        }
    }

    public async updateTask(bodyId: string, taskId: string, termId: string, attendanceFormatId: string, comment: string): promise<IUserTaskResponse[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/item/' + taskId + '/option/' + termId + '/attendanceFormat/' + attendanceFormatId + '/comment', AadHttpClient.configurations.v1, { body: JSON.stringify(comment), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to update the task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async createDelegateTask(bodyId: string, taskDelegate: IUserTaskDelegateModel): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/delegatetask', AadHttpClient.configurations.v1, { body: JSON.stringify(taskDelegate) });
            await BackendService.handleServerSideException(response);
            // return await response.json();

        } catch (error) {
            Logger.error('Failed to created delegate task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateTaskDelegate(bodyId: string, taskId: string, termId: string, comment: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/item/' + taskId + '/optionDelegated/' + termId + '/comment', AadHttpClient.configurations.v1, { body: JSON.stringify(comment), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            // return await response.json();
        } catch (error) {
            Logger.error('Failed to update the task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateMinutesInformation(bodyId: string, eventId: string, minutes: IBackMinutes): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/minutesinformation', AadHttpClient.configurations.v1, { body: JSON.stringify(minutes), method: 'POST' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to update the minutes from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async approdvalTaskMinutes(bodyId: string, taskId: string, termId: string, comment: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/item/' + taskId + '/optionApprodvalminutes/' + termId + '/comment', AadHttpClient.configurations.v1, { body: JSON.stringify(comment), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to update the task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async modificationTaskMinutes(bodyId: string, taskId: string, termId: string, comment: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/item/' + taskId + '/optionModificationminutes/' + termId + '/comment', AadHttpClient.configurations.v1, { body: JSON.stringify(comment), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to update the task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async startTheMinuteApprodvalprocess(bodyId: string, eventId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + '/SendRequestApprodvalsMinutes', AadHttpClient.configurations.v1, { method: 'POST' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to start the minute approdval process from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async requestCertificate(bodyId: string, taskCertification: ITaskCertification): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/certificationtask', AadHttpClient.configurations.v1, { body: JSON.stringify(taskCertification), method: 'POST' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to request the certificate from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async approdvalTaskCertification(bodyId: string, taskId: string, termId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/item/' + taskId + '/optionCertification/' + termId, AadHttpClient.configurations.v1, { method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error('Failed to approdval the task certification from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getApprodvalStatusMinutes(bodyId: string, eventId: string): promise<IApprodvalStatusOfTheMinutes> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/events/' + eventId + '/statusapprodvalminutes', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get the approdval status of the minutes from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async downloadAgreementsTemplate(bodyId: string, eventId: string, agreementId: string): promise<ArrayBuffer> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/agreements/" + agreementId + "/downloadTemplate", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.arrayBuffer();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async downloadOrderOfTheDayTemplate(bodyId: string, eventId: string): promise<ArrayBuffer> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/agendaItems/downloadTemplate", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.arrayBuffer();
        } catch (error) {
            Logger.error('Failed to download template of the order of the day from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async downloadMinutesTemplate(bodyId: string, eventId: string): promise<ArrayBuffer> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/minutes/downloadTemplate", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.arrayBuffer();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async downloadAttendanceTemplate(bodyId: string, eventId: string): promise<ArrayBuffer> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/attendance/downloadTemplate", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.arrayBuffer();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getEventVotationDetail(bodyId: string, eventId: string, fromArchive?: boolean): promise<IEventVotationDetails[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/votationDetail" + (fromArchive ? "?fromArchive=true" : ""), AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }
    public async addOrUpdateEventVotationDetail(bodyId: string, eventId: string, votations: IEventVotationDetails[]): promise<IEventVotationDetails[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/votationDetail", AadHttpClient.configurations.v1, { body: JSON.stringify(votations), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }
    public async getEventVotingRetestsentation(bodyId: string, eventId: string, fromArchive?: boolean): promise<Record<string, string[]>> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/votingRetestsentation", AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }
    public async startEventVotation(bodyId: string, eventId: string): promise<IEventVotationDetails[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/startVotation", AadHttpClient.configurations.v1, { method: 'POST' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }
    public async updateVoteByUser(bodyId: string, eventId: string, vote: IVote): promise<IEventVotationDetails[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/updateVote", AadHttpClient.configurations.v1, { body: JSON.stringify(vote), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateVoteByRetestsentation(bodyId: string, eventId: string,retestsentationId: string, vote: IVote): promise<IEventVotationDetails> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/bodies/' + bodyId + '/events/' + eventId + "/updateVoteByRetestsentation/" + retestsentationId, AadHttpClient.configurations.v1, { body: JSON.stringify(vote), method: 'PUT' });
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to download template agreement from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getOrganInfo(bodyId: string): promise<IdepartmentManagement[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/bodies', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get organ information:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getUsersInfo(bodyId: string): promise<Record<string, IUserManagement>> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/users', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get users information:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async getUserTaskByTaskId(taskId:string, bodyId: string): promise<IUserTaskModel> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/tasks/bodies/' + bodyId + '/taskId/' + taskId, AadHttpClient.configurations.v1,);
            await BackendService.handleServerSideException(response);
            return await response.json();
        } catch (error) {
            Logger.error('Failed to get user task from bodyId:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateUserprofile(userInfo: IUserManagement, bodyId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/user', AadHttpClient.configurations.v1, { body: JSON.stringify(userInfo), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async updateUsersInOrgan(membersInfo: IUsersOrganRelation[], bodyId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/body/users', AadHttpClient.configurations.v1, { body: JSON.stringify(membersInfo), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async updateOrganInfo(organInfo: IdepartmentManagement, bodyId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/body', AadHttpClient.configurations.v1, { body: JSON.stringify(organInfo), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to get user profile`, error, this._pageContext);
            throw error;
        }
    }

    public async checkUserHasPermissions(checkInfo: IUsersOrganRelation): promise<IPermissionCheckResult> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/management/permissions', AadHttpClient.configurations.v1, { body: JSON.stringify(checkInfo) });
            await BackendService.handleServerSideException(response);
            const res: IPermissionCheckResult = await response.json();
            return res;
        } catch (error) {
            Logger.error(`Failed to check user Permissions from admin app`, error, this._pageContext);
            throw error;
        }
    }

    public async getEXTERNALOrgansInfo(bodyId: string): promise<IOrganEXTERNAL[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/EXTERNALBodies', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get EXTERNAL organs information:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateEXTERNALOrganInfo(organInfo: IOrganEXTERNAL, bodyId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/EXTERNALBody', AadHttpClient.configurations.v1, { body: JSON.stringify(organInfo), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to update EXTERNAL organ info from bodyId: ${bodyId}`, error, this._pageContext);
            throw error;
        }
    }

    public async createEXTERNALOrgan(organInfo: IOrganEXTERNAL, bodyId: string): promise<void> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/management/' + bodyId + '/EXTERNALBody', AadHttpClient.configurations.v1, { body: JSON.stringify(organInfo)});
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to create new EXTERNAL organ from bodyId: ${bodyId}`, error, this._pageContext);
            throw error;
        }
    }
    
    public async downloadBodyCertificate(bodyId: string,body:string, isEXTERNAL?:boolean): promise<ArrayBuffer> {
        const client = await this._aadHttpClient;
        try {
            
            if(isEXTERNAL){
                const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/EXTERNALBody/downloadTemplate/'+body, AadHttpClient.configurations.v1);
                await BackendService.handleServerSideException(response);
                return await response.arrayBuffer();
            }else{
                const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/body/downloadTemplate/'+body, AadHttpClient.configurations.v1);
                await BackendService.handleServerSideException(response);
                return await response.arrayBuffer();
            }
            
           /* return new promise((resolve) => {
                setTimeout(() => {
                    const eCNTyArrayBuffer = new ArrayBuffer(0);
                    resolve(eCNTyArrayBuffer);
                  }, 3000); // Espera de 3 segundos (3000 milisegundos)
              });*/
        } catch (error) {
            Logger.error('Failed to download organ certificate for organ:' + body, error, this._pageContext);
            throw error;
        }
    }

    public async getAreaRelationsInfo(bodyId: string): promise<IRelationBusinessArea[]> {
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/BusinessUnit', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get area sectorial - Division information:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async updateAreaRelationsInfo(item:IRelationBusinessArea, bodyId: string):promise<void>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/BusinessUnit', AadHttpClient.configurations.v1, { body: JSON.stringify(item), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to update area relations info`, error, this._pageContext);
            throw error;
        }
    }

    public async getAvailableLicenses(bodyId:string):promise<ILicenseInformation[]>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/management/' + bodyId + '/availableLicenses', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get licenses information:' + bodyId, error, this._pageContext);
            throw error;
        }
    }

    public async manageUserLicense(bodyId:string, upn:string, license:string):promise<void>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/management/' + bodyId + '/manageUserLicense', AadHttpClient.configurations.v1, 
            { 
                body: JSON.stringify({
                    UserPrincipalName:upn, 
                    AssignedLicenses:license ? [license]: []
            }), method: 'PUT' });
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to update user license: ${upn} + ${license}`, error, this._pageContext);
        }
    }

    public async createUser(item:BodyUserInformation, bodyId: string):promise<void>{
        console.log(JSON.stringify(item));
        const client = await this._aadHttpClient;
        try {
            const response = await client.post(EnvConfig.BackendHost + '/api/management/' + bodyId + '/user', AadHttpClient.configurations.v1, { body: JSON.stringify(item)});
            await BackendService.handleServerSideException(response);
        } catch (error) {
            Logger.error(`Failed to create new user`, error, this._pageContext);
            throw error;
        }
    }

    public async getSensitivityLabels():promise<ITenantSensitivityLabel[]>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/sensitivitylabels', AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get sensitivityLabels:', error, this._pageContext);
            throw error;
        }
    }

    public async getDocumentsSensitivityLabelsByEventSharedId(bodyId:string, eventId: string, fromArchive?:boolean, isEditor?:boolean):promise<IDocumentSensitivityLabel[]>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.get(EnvConfig.BackendHost + '/api/bodies/'+bodyId+'/events/'+eventId+'/documentSensitivityLabels?fromArchive='+(fromArchive ? "true" : "false")+'&isEditor='+(isEditor ? "true": "false"), AadHttpClient.configurations.v1);
            
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to get DocumentsSensitivityLabels for bodyId: '+bodyId+' and event: '+eventId, error, this._pageContext);
            throw error;
        }
    }

    public async updateSensitivityLabel(documentSensitivityLabel:IDocumentSensitivityLabel, fromArchive?:boolean, isEditor?:boolean):promise<boolean>{
        const client = await this._aadHttpClient;
        try {
            const response = await client.fetch(EnvConfig.BackendHost + '/api/sensitivitylabels?fromArchive='+(fromArchive ? "true" : "false")+'&isEditor='+(isEditor ? "true": "false"), AadHttpClient.configurations.v1,{ method: 'PATCH', body: JSON.stringify(documentSensitivityLabel)});
            await BackendService.handleServerSideException(response);
            //return await response.json();
            const res = await response.json();
            return res;
        } catch (error) {
            Logger.error('Failed to update DocumentsSensitivityLabels', error, this._pageContext);
            throw error;
        }
    }
    public async launchSaveDataprocess(startDate?:string): promise<string> {
        const client = await this._aadHttpClient;
        try {
            ///api/powerbireports/{startDate?}
            const response = await client.post(EnvConfig.BackendHost + (startDate ? '/api/powerbireports/'+startDate : '/api/powerbireports') , AadHttpClient.configurations.v1, { method: 'POST' });
            await BackendService.handleServerSideException(response);
            return await response.json();
            //console.log(startDate);
            //return startDate ? "llamando back completa":"llamando back incremental"
        } catch (error) {
            Logger.error('Failed to launch powerBi report process', error, this._pageContext);
            throw error;
        }
    }

    public async getUserInvitationMessage(serverRelativeUrl:string): promise<string> {
        const client = await this._aadHttpClient;
        try {
            const encodedUrl = encodeURIComponent(serverRelativeUrl);
            const response = await client.get(EnvConfig.BackendHost + '/api/management/invitationText?serverRelativeUrl='+encodedUrl , AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.text();
        } catch (error) {
            Logger.error('Failed to get new user invitations', error, this._pageContext);
            throw error;
        }
    }
    
    public async getAllEventsFromBbdd(startDate:string, endDate:string): promise<IEventFromBbdd[]> {
        const client = await this._aadHttpClient;
        try {
            ///api/powerbireports/{startDate?}
            const response = await client.get(EnvConfig.BackendHost + '/api/events/all?startDate=' +startDate+ '&endDate='+endDate+'&fromArchive=false' , AadHttpClient.configurations.v1);
            await BackendService.handleServerSideException(response);
            return await response.json();
            //console.log(startDate);
            //return startDate ? "llamando back completa":"llamando back incremental"
        } catch (error) {
            Logger.error('Failed to get All events from BBDD', error, this._pageContext);
            throw error;
        }
    }
/*
    public async getAllEventsFromBbdd(startDate: string, endDate: string): promise<IEventFromBbdd[]> {
        return new promise((resolve) => {
            setTimeout(() => {
                resolve([
                    {
                        RowKey: "5169aa50bae34b6aad830f565154ef47",
                        EstadoMeeting: "test reserva",
                        EndDate: new Date("2025-09-09T22:17:00Z"),
                        StartDate: new Date("2025-09-09T21:17:39Z"),
                        NombreDepartment: "Comisión Nacional de Astronomía",
                        Department: "demoev-co-cna",
                        TipoDepartment: "Comisión",
                        MeetingType: "Online",
                        Title: "Test Meeting test-Reserva"
                    },
                    {
                        RowKey: "8fc0a4694bf64bee8ff50898cba6bdb7",
                        EstadoMeeting: "Celebrada",
                        EndDate: new Date("2025-09-09T22:21:00Z"),
                        StartDate: new Date("2025-09-09T21:21:38Z"),
                        NombreDepartment: "Comisión Nacional de Astronomía",
                        Department: "demoev-co-cna",
                        TipoDepartment: "Comisión",
                        MeetingType: "Online",
                        Title: "Test Meeting Publicada"
                    }
                ]);
            }, 3000); // 3 segundos
        });
    }*/    
}