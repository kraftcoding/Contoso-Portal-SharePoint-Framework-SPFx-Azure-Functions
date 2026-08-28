using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using PnP.Core;
using PnP.Core.Services;
using System.Net;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Http.EventsHttpTrigger
{
    public class EventsHttpTrigger : TriggerBase<EventsHttpTrigger>
    {
        #region 
        private NotificationsService _notificationService;
        private EventsService _eventsService;
        private EventAttendanceService _eventAttendanceService;
        private EventAttendanceTemplateService _eventAttendanceTemplateService;
        private EventCaptureService _eventCaptureService;
        private EventPublishingService _eventPublishingService;
        private EventMinutesService _eventMinutesService;
        private EventMinutesTemplateService _eventMinutesTemplateService;
        private EventAgendaItemsService _eventAgendaItemsService;
        private EventAgreementService _eventAgreementService;
        private EventVotationService _eventVotationService;
        private EventFilesSensitivityLabelsService _eventFilesSensitivityLabelsService;

        public EventsHttpTrigger(
            EventsService eventsService,
            EventAttendanceService eventAttendanceService,
            EventAttendanceTemplateService eventAttendanceTemplateService,
            EventCaptureService eventCaptureService,
            EventPublishingService eventPublishingService,
            EventAgreementService eventAgreementService,
            EventMinutesService eventMinutesService,
            EventMinutesTemplateService eventMinutesTemplateService,
            EventAgendaItemsService eventAgendaItemsService,
            EventVotationService eventVotationService,
            EventFilesSensitivityLabelsService eventFilesSensitivityLabelsService,
            NotificationsService notificationsService,
            ILogger<EventsHttpTrigger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            this._eventsService = eventsService;
            this._eventAttendanceService = eventAttendanceService;
            _eventAttendanceTemplateService = eventAttendanceTemplateService;
            this._eventCaptureService = eventCaptureService;
            this._eventPublishingService = eventPublishingService;
            this._eventAgreementService = eventAgreementService;
            this._eventMinutesService = eventMinutesService;
            _eventMinutesTemplateService = eventMinutesTemplateService;
            _eventAgendaItemsService = eventAgendaItemsService;
            this._eventVotationService = eventVotationService;
            this._eventFilesSensitivityLabelsService = eventFilesSensitivityLabelsService;
            this._notificationService = notificationsService;
        }

        #endregion

        #region Managed mails for events

        [Function(nameof(GetManagedMails))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_managed_mails", tags: new[] { "events" }, Summary = "Return mails for management.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(List<Object>), Summary = "")]
        public async Task<HtttestsponseData> GetManagedMails([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/managed-mails/{eventId:required}")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

            Dictionary<string, EmailTypeContent> mails = new();

            var eventDAO = await this._eventsService.GetBySharedEventId(ctx, bodyId, eventId, fromArchive: false);

            //VAMOS A DEVOLVER SOLAMENTE ESTOS TIPOS DE MENSAJE: PUBLICAR, EN CELEBRACIÓN Y CANCELADA, EN FUNCIÓN DEL ESTADO DE LA Meeting

            if (eventDAO == null)
                throw new Exception("Event not found");

            List<string> availableMailTypes = new List<string>();

            if (eventDAO.StatusId == TaxonomyValuesIds.EventStatus.InConstruction)
            {
                availableMailTypes.Add("Publicar");
                availableMailTypes.Add("EnCelebracion");
                availableMailTypes.Add("Cancelada");
            }
            if (eventDAO.StatusId == TaxonomyValuesIds.EventStatus.Published)
            {
                availableMailTypes.Add("EnCelebracion");
                availableMailTypes.Add("Cancelada");
            }

            //NotificationMessagesJsonDAO messagesJsonDAO = await _notificationService.GetMessagesAll();
            ConfigEmailBodiesJsonDAO bodyMails = await this._eventsService.GetBodyAllMails(ctx, bodyId, eventId);
            string managedMailsJson = await this._eventsService.GetEmailJsonBySharedEventId(ctx, bodyId, eventId);

            foreach (var template in availableMailTypes)
            {
                switch (template)
                {
                    case "Publicar":
                        if (managedMailsJson != null)
                        {
                            var managedMails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(managedMailsJson);
                            if (managedMails != null && managedMails.ContainsKey("Publicar"))
                            {
                                mails.Add(template, managedMails["Publicar"]);
                                break;
                            }
                        }
                        mails.Add(template, bodyMails.BodyTypes.Published);
                        break;
                    case "EnCelebracion":
                        if (managedMailsJson != null)
                        {
                            var managedMails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(managedMailsJson);
                            if (managedMails != null && managedMails.ContainsKey("EnCelebracion"))
                            {
                                mails.Add(template, managedMails["EnCelebracion"]);
                                break;
                            }
                        }
                        mails.Add(template, bodyMails.BodyTypes.InCelebration);
                        break;
                    case "Cancelada":
                        if (managedMailsJson != null)
                        {
                            var managedMails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(managedMailsJson);
                            if (managedMails != null && managedMails.ContainsKey("Cancelada"))
                            {
                                mails.Add(template, managedMails["Cancelada"]);
                                break;
                            }
                        }
                        mails.Add(template, bodyMails.BodyTypes.Cancelled);
                        break;
                }
            }

            var settings = new JsonSerializerSettings
            {
                StringEscapeHandling = StringEscapeHandling.EscapeHtml
            };

            string mailsAllJson = JsonConvert.SerializeObject(mails, Formatting.Indented, settings);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(mailsAllJson).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(SaveManagedMails))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "save_managed_mails", tags: new[] { "events" }, Summary = "Save mails for management.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(Dictionary<string, EmailTypeContent>), Required = true, Description = "Mails to save for management.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Dictionary<string, EmailTypeContent>), Summary = "Returns the saved mails.")]
        public async Task<HtttestsponseData> SaveManagedMails([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/managed-mails/{eventId:required}")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var mails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(requestBody);

            // 2. Logic to save Email.json inside the event's Document Set
            var saveResult = await this._eventsService.SaveEmailJsonBySharedEventId(ctx, bodyId, eventId, JsonConvert.SerializeObject(mails));

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json; charset=utf-8");
            await response.WriteStringAsync(JsonConvert.SerializeObject(mails));

            return response;
        }

        [Function(nameof(SendpreviewManagedMails))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "send_preview_managed_mails", tags: new[] { "events" }, Summary = "Send preview email for a managed mail template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string))]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string))]
        [OpenApiParameter(name: "emailId", In = ParameterLocation.Path, Required = true, Type = typeof(string))]
        [OpenApiParameter(name: "eventName", In = ParameterLocation.Path, Required = true, Type = typeof(string))]
        [OpenApiParameter(name: "mailContent", In = ParameterLocation.Path, Required = true, Type = typeof(string))]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(object))]
        public async Task<HtttestsponseData> SendpreviewManagedMails([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/managed-mails/{eventId:required}/send-preview/{emailId:required}/{eventName:required}")] HtttestquestData req, string bodyId, string eventId, string emailId, string eventName)
        {
            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            using var rootCtx = await CreatePnPContextAsSystem();

            // 1. Read JSON body
            var requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var mails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(requestBody);

            string bodypreview = mails.GetValueOrDefault(eventName)?.Body?"en-US";
            string subjectpreview = eventName + " (testvisualización)";

            if (string.IsNullOrECNTy(bodypreview))
                bodypreview = "Error en la plantilla de correo. ContMinutes con el administrador.";

            var currentUser = await ctx.Web.GetCurrentUserAsync();

            await _notificationService.SendEmailNotificationpreview(rootCtx, bodypreview, subjectpreview, currentUser.Mail);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json; charset=utf-8");

            return response;
        }

        #endregion

        #region Event

        [Function(nameof(CreateEvent))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "create_events", tags: new[] { "events" }, Summary = "Create a new event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(NewOrUpdatedEvent), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event), Summary = "")]
        public async Task<HtttestsponseData> CreateEvent([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events")] HtttestquestData req, string bodyId)
        {
            var theEvent = JsonConvert.DeserializeObject<NewOrUpdatedEvent>(await new StreamReader(req.Body).ReadToEndAsync());
            if (theEvent == null)
                throw new ArgumentException("Invalid request body");

            // create context for the body site 
            using var ctx = await CreatePnPContextAsUser(req);

            var newEvent = await _eventsService.CreateEvent(ctx, bodyId, theEvent);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(newEvent)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateEventById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_event", tags: new[] { "events" }, Summary = "Update an event by ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(NewOrUpdatedEvent), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event), Summary = "")]
        public async Task<HtttestsponseData> UpdateEventById([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}")] HtttestquestData req, string bodyId, string eventId)
        {
            var theEvent = JsonConvert.DeserializeObject<NewOrUpdatedEvent>(await new StreamReader(req.Body).ReadToEndAsync());
            if (theEvent == null)
                throw new ArgumentException("Invalid request body");

            using var ctx = await CreatePnPContextAsUser(req);

            var updatedEvent = await _eventsService.UpdateBySharedEventId(ctx, bodyId, eventId, theEvent);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(updatedEvent)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(SendRequestApprodvalsMinutes))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "SendRequestApprodvalsMinutes", tags: new[] { "events" }, Summary = "Send approdval request for minutes.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventSendOperationInformation), Summary = "")]
        public async Task<HtttestsponseData> SendRequestApprodvalsMinutes([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/SendRequestApprodvalsMinutes")] HtttestquestData req, string bodyId, string eventId)
        {
            HtttestsponseData response;
            using var ctx = await CreatePnPContextAsUser(req);

            await _eventsService.SendRequestApprodvalsMinutes(ctx, bodyId, eventId);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetEventById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event", tags: new[] { "events" }, Summary = "Return event details by ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event), Summary = "")]
        public async Task<HtttestsponseData> GetEventById([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var events = await _eventsService.GetBySharedEventId(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(events)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetUserEvents))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_events", tags: new[] { "events" }, Summary = "Return all events of the current user.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the events are from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserEvents([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "events")] HtttestquestData req)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var events = await _eventsService.GetUserEvents(ctx, bodyId: null, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(events)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetAllEvents))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_events", tags: new[] { "events" }, Summary = "Return all events of the current platform.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the events are from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event[]), Summary = "")]
        public async Task<HtttestsponseData> GetAllEvents([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "events/all")] HtttestquestData req)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            _ = DateTime.TryParse(req.Query["startDate"], out var startDate);
            _ = DateTime.TryParse(req.Query["endDate"], out var endDate);
            var events = await _eventsService.GetAllEventsFromDataStorageByDate(ctx, startDate, endDate, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(events)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetUserEventsByBody))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_events_bybody", tags: new[] { "events" }, Summary = " Return all events of the current user by body.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the events are from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Event[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserEventsByBody([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events")] HtttestquestData req, string bodyId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var events = await _eventsService.GetUserEvents(ctx, bodyId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(events)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(ChangeEventStatusTo))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "changestatus_event_byid", tags: new[] { "events" }, Summary = "Change the status of an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "changeMessage", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Reason for event changes.")]
        [OpenApiParameter(name: "statusId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the current status.")]
        [OpenApiParameter(name: "skipUserCommunications", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Indicate to skip user communications.")]
        [OpenApiParameter(name: "forceStateChange", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventSendOperationInformation), Summary = "")]
        public async Task<HtttestsponseData> ChangeEventStatusTo([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/changeStatus")] HtttestquestData req, string bodyId, string eventId, [DurableClient] DurableTaskClient client)
        {
            var changeMessage = req.Query["changeMessage"];
            var newStatus = req.Query["statusId"];
            if (string.IsNullOrECNTy(newStatus))
                throw new ArgumentException($"Parameter required {newStatus} is not provided");

            _ = bool.TryParse(req.Query["skipUserCommunications"], out var skipUserCommunications);
            _ = bool.TryParse(req.Query["forceStateChange"], out var forceStateChange);

            using var ctx = await CreatePnPContextAsUser(req);

            await _eventsService.ChangeStatusByEvent(ctx, bodyId, eventId, newStatus);
            var (instanceId, operation) = await SendEventChanges(ctx, client, bodyId, eventId, changeMessage, new EventSendOperationSettings() { SkipUserCommunications = skipUserCommunications, ForceStateChange = forceStateChange });

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(operation)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(SendEventUpdatesById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "send_event", tags: new[] { "events" }, Summary = "Send updates for an event by ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "changeMessage", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Reason for event changes.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventSendOperationInformation), Summary = "")]
        public async Task<HtttestsponseData> SendEventUpdatesById([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/send")] HtttestquestData req, string bodyId, string eventId, [DurableClient] DurableTaskClient client)
        {
            var changeMessage = req.Query["changeMessage"];
            using var ctx = await CreatePnPContextAsUser(req);

            var (instanceId, operation) = await SendEventChanges(ctx, client, bodyId, eventId, changeMessage);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(operation)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(EventPendingChanges))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_pendingChanges", tags: new[] { "events" }, Summary = "Return pending changes for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(PendingChanges[]), Summary = "")]
        public async Task<HtttestsponseData> EventPendingChanges([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/pendingChanges")] HtttestquestData req, string bodyId, string eventId, [DurableClient] DurableTaskClient client)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var response = req.CreateResponse(HttpStatusCode.OK);
            var hasPendingChanges = false;

            var isprocessing = await OrchestrationHelpers.IsPendingOrchestrations(client, bodyId, eventId);
            List<PendingChanges>? pendingChanges = [];
            if (!isprocessing)
            {
                pendingChanges = await _eventsService.HasPendingChanges(ctx, bodyId, eventId);
                hasPendingChanges = pendingChanges is not null;
            }

            await response.WriteStringAsync(JsonConvert.SerializeObject(pendingChanges)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }


        private async Task<(string InstanceId, EventSendOperationInformation Operation)> SendEventChanges(IPnPContext userCtx, DurableTaskClient client, string bodyId, string eventId, string changeMessage, EventSendOperationSettings? settings = null)
        {
            var operation = await _eventCaptureService.CaptureAndprepareForSendingByEvent(userCtx, bodyId, eventId, changeMessage, settings);
            var instanceId = await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableEventSendingOrchestration), operation, new StartOrchestrationOptions() { InstanceId = $"{OrchestrationHelpers.GetInstanceIdForEventChanges(bodyId, eventId)}-c:{operation.ToCaptureState.CaptureId}" });

            return (instanceId, operation);
        }

        #endregion

        #region Event attendance

        [Function(nameof(AddOrUpdateAttendance))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_attendanceList", tags: new[] { "events" }, Summary = "Add or update attendance for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "attendance", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "List of event attendees.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventUserAttendance[]), Summary = "")]
        public async Task<HtttestsponseData> AddOrUpdateAttendance([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/addAttendance/{attendance:required}")] HtttestquestData req, string bodyId, string eventId, string attendance)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var usersAttendance = await _eventAttendanceService.AddOrRemoveEventAttendee(ctx, bodyId, eventId, attendance, true);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(usersAttendance)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DeleteAttendance))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_attendanceList", tags: new[] { "events" }, Summary = "Remove attendance from an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "attendance", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "List of event attendees.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventUserAttendance[]), Summary = "")]
        public async Task<HtttestsponseData> DeleteAttendance([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/deleteAttendance/{attendance:required}")] HtttestquestData req, string bodyId, string eventId, string attendance)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var usersAttendance = await _eventAttendanceService.AddOrRemoveEventAttendee(ctx, bodyId, eventId, attendance, false);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(usersAttendance)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(UpdateAttendance))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_attendanceList", tags: new[] { "events" }, Summary = "Update attendance for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateAttendance([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/updateAttendance")] HtttestquestData req, string bodyId, string eventId)
        {
            var attendances = JsonConvert.DeserializeObject<EventUserAttendance[]>(await new StreamReader(req.Body).ReadToEndAsync());
            if (attendances == null)
                throw new ArgumentException("Invalid request body");
            using var ctx = await CreatePnPContextAsUser(req);

            var usersAttendance = await _eventAttendanceService.UpdateAttendance(ctx, bodyId, eventId, attendances.ToList());

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        #endregion

        #region Event agenda items

        [Function(nameof(GetEventAgendaItems))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_agendaitems", tags: new[] { "events" }, Summary = "Return agenda items for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventAgendaItem[]), Summary = "")]
        public async Task<HtttestsponseData> GetEventAgendaItems([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/items")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var agendaItems = await this._eventAgendaItemsService.GetAgendaItemsByEvent(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agendaItems)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(AddOrUpdateEventAgendaItems))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_agendaitem", tags: new[] { "events" }, Summary = "Add or update agenda items for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(EventAgendaItem[]), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventAgendaItem[]), Summary = "")]
        public async Task<HtttestsponseData> AddOrUpdateEventAgendaItems([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/items")] HtttestquestData req, string bodyId, string eventId)
        {
            var agendaItemToUpdate = JsonConvert.DeserializeObject<EventAgendaItem[]>(await new StreamReader(req.Body).ReadToEndAsync()) ?? throw new ArgumentException("Invalid request body");
            using var ctx = await CreatePnPContextAsUser(req);

            var agendaItems = await this._eventAgendaItemsService.AddOrUpdateAgendaItemsForEvent(ctx, bodyId, eventId, agendaItemToUpdate);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agendaItems)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DeleteEventAgendaItems))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "delete_event_agendaitem", tags: new[] { "events" }, Summary = "Remove agenda items from an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agendaItemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agenda item.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.Accepted, Summary = "")]
        public async Task<HtttestsponseData> DeleteEventAgendaItems([HttpTrigger(AuthorizationLevel.Anonymous, "DELETE", Route = "bodies/{bodyId:required}/events/{eventId:required}/items/{agendaItemId:required}")] HtttestquestData req, string bodyId, string eventId, string agendaItemId)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventAgendaItemsService.DeleteAgendaItem(ctx, bodyId, eventId, agendaItemId);

            var response = req.CreateResponse(HttpStatusCode.Accepted);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetUserAttendancesByEventId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_attendance_users", tags: new[] { "events" }, Summary = "Return user attendances by event ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventUserAttendance[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserAttendancesByEventId([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/attendance")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var attendance = await this._eventAttendanceService.GetEventAttendance(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(attendance)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetEventAgreementItems))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_agreements", tags: new[] { "events" }, Summary = "Return event agreement items.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventAgreement[]), Summary = "")]
        public async Task<HtttestsponseData> GetEventAgreementItems([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/agreements")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var agreements = await this._eventAgreementService.GetAgreementsByEvent(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agreements)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetAgreementIdsWithPendingCertTask))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_agreements", tags: new[] { "events" }, Summary = "Return agreement ids with pending certificate tasks.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string[]), Summary = "")]
        public async Task<HtttestsponseData> GetAgreementIdsWithPendingCertTask([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/pendingcertagreements")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            var agreementsIdsWithPendingCertTask = await this._eventAgreementService.GetAgreementIdsWithPendingCertTask(ctx, bodyId, eventId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agreementsIdsWithPendingCertTask)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(AddOrUpdateAgreements))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_agreements", tags: new[] { "events" }, Summary = "Add or update agreements for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(EventAgreement[]), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventAgreement[]), Summary = "")]
        public async Task<HtttestsponseData> AddOrUpdateAgreements([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/agreements")] HtttestquestData req, string bodyId, string eventId, [DurableClient] DurableTaskClient client)
        {
            var agreementsToUpdate = JsonConvert.DeserializeObject<EventAgreement[]>(await new StreamReader(req.Body).ReadToEndAsync());
            if (agreementsToUpdate == null)
                throw new ArgumentException("Invalid request body");

            using var ctx = await CreatePnPContextAsUser(req);

            var agreements = await this._eventAgreementService.AddOrUpdateAgreementsForEvent(ctx, bodyId, eventId, agreementsToUpdate);
            var (instanceId, operation) = await SendEventChanges(ctx, client, bodyId, eventId, "Actualización de voto");

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agreements)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DownloadAttendanceTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_attendance_template", tags: new[] { "events" }, Summary = "Download the attendance template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadAttendanceTemplate(
                        [HttpTrigger(AuthorizationLevel.Anonymous, "GET",
                            Route = "bodies/{bodyId:required}/events/{eventId:required}/attendance/downloadTemplate")] HtttestquestData req,
                        string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            var attendanceTemplate = await _eventAttendanceTemplateService.DownloadAttendaceTemplate(ctx, bodyId, eventId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/pdf"); // .pdf
            response.Headers.Add("Content-Disposition", "attachment; filename=filename.pdf"); // Uncomment to download in swagger
            await response.WriteBytesAsync(attendanceTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        #endregion

        #region Event documents

        [Function(nameof(AddRelatedDocumentByUniqueId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addrelated_doc_agendaitem", tags: new[] { "events" }, Summary = "Associate a document by its ID to an agenda item.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agendaItemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agenda item.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> AddRelatedDocumentByUniqueId([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/items/{agendaItemId:required}/addRelatedDocumentByUniqueId")] HtttestquestData req, string bodyId, string eventId, string agendaItemId)
        {
            var docId = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(docId))
                throw new ArgumentException($"Parameter required {docId} is not implemented");

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventAgendaItemsService.AddDocumentToAgendaItemByUniqueId(ctx, bodyId, eventId, agendaItemId, docId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(RemoveRelatedDocumentByUniqueId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "removerelated_doc_agendaitem", tags: new[] { "events" }, Summary = "Remove the association of a document by its ID from an agenda item.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agendaItemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agenda item.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> RemoveRelatedDocumentByUniqueId([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/items/{agendaItemId:required}/removeRelatedDocumentByUniqueId")] HtttestquestData req, string bodyId, string eventId, string agendaItemId)
        {
            var docId = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(docId))
                throw new ArgumentException($"Parameter required {docId} is not implemented");

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventAgendaItemsService.RemoveDocumentToAgendaItemByUniqueId(ctx, bodyId, eventId, agendaItemId, docId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(AddRelatedDocumentIdToAgreement))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addrelated_doc_agreementitem", tags: new[] { "events" }, Summary = "Associate a document by its ID with an agreement.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agreementItemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agreement.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiParameter(name: "isCertificate", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = " Indicates if the document is a certificate.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> AddRelatedDocumentIdToAgreement([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/items/{agreementItemId:required}/addRelatedDocumentIdToAgreement")] HtttestquestData req, string bodyId, string eventId, string agreementItemId)
        {
            var docId = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(docId))
                throw new ArgumentException($"Parameter required {docId} is not implemented");

            var isCertificate = req.Query["isCertificate"];

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventAgreementService.AddRemoveDocumentToAgreementByUniqueId(ctx, bodyId, eventId, agreementItemId, docId, Convert.ToBoolean(isCertificate), true);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(RemoveRelatedDocumentIdToAgreement))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "removerelated_doc_agreementitem", tags: new[] { "events" }, Summary = "Remove the association of a document by its ID from an agreement.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agreementItemId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agreement.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> RemoveRelatedDocumentIdToAgreement([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/items/{agreementItemId:required}/removeRelatedDocumentIdToAgreement")] HtttestquestData req, string bodyId, string eventId, string agreementItemId)
        {
            var docId = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(docId))
                throw new ArgumentException($"Parameter required {docId} is not implemented");

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventAgreementService.AddRemoveDocumentToAgreementByUniqueId(ctx, bodyId, eventId, agreementItemId, docId, false, false);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetEventStorage))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_storage", tags: new[] { "events" }, Summary = "Return storage information for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventsService.EventStorage), Summary = "")]
        public async Task<HtttestsponseData> GetEventStorage([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/docs")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var storage = await this._eventsService.GetEventDocumentStorages(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(storage)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DeleteDocumentByUniqueId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "delete_doc", tags: new[] { "events" }, Summary = "Remove a document by its ID.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DeleteDocumentByUniqueId([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/deleteDocumentByUniqueId")] HtttestquestData req, string bodyId, string eventId)
        {
            var docId = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(docId))
                throw new ArgumentException($"Parameter required {docId} is not implemented");

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventsService.DeleteDocumentByUniqueId(ctx, bodyId, eventId, docId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        #endregion

        #region Minutes Service

        [Function(nameof(GetMinutesInformation))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_minutes_information", tags: new[] { "events" }, Summary = "Return the information of the minutes.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventMinutesInformation), Summary = "")]
        public async Task<HtttestsponseData> GetMinutesInformation([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/minutesinformation")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var storage = await this._eventMinutesService.GetMinutesInformationByEvent(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(storage)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateMinutesInformation))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_minutes_information", tags: new[] { "events" }, Summary = "Update the information of the minutes.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(EventMinutesInformation), Required = true)]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> UpdateMinutesInformation([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/minutesinformation")] HtttestquestData req, string bodyId, string eventId)
        {
            HtttestsponseData response;
            using var ctx = await CreatePnPContextAsUser(req);
            var minutesInformation = JsonConvert.DeserializeObject<EventMinutesInformation>(await new StreamReader(req.Body).ReadToEndAsync());

            await this._eventMinutesService.SetMinutesInformation(ctx, bodyId, eventId, minutesInformation);

            response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(SetMinutesDocument))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "set_minutes_document", tags: new[] { "events" }, Summary = "Associate a document with the minutes.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "docUniqueId", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "ID of the document.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> SetMinutesDocument([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/SetMinutesDocument")] HtttestquestData req, string bodyId, string eventId, [DurableClient] DurableTaskClient client)
        {
            var uIdDocument = req.Query["docUniqueId"];
            if (string.IsNullOrECNTy(uIdDocument))
                throw new ArgumentException($"Parameter required {uIdDocument} is not implemented");

            using var ctx = await CreatePnPContextAsUser(req);

            await this._eventMinutesService.SetMinutesDocument(ctx, bodyId, eventId, uIdDocument);

            var (instanceId, operation) = await SendEventChanges(ctx, client, bodyId, eventId, "Asociación de Minutes: publicación automática");

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DownloadMinutesTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_minutes_template", tags: new[] { "events" }, Summary = "Download the minutes template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadMinutesTemplate(
                        [HttpTrigger(AuthorizationLevel.Anonymous, "GET",
                            Route = "bodies/{bodyId:required}/events/{eventId:required}/minutes/downloadTemplate")] HtttestquestData req,
                        string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var minutesTemplate = await _eventMinutesTemplateService.DownloadMinutesTemplate(ctx, bodyId, eventId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"); // .docx
            await response.WriteBytesAsync(minutesTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        #endregion

        #region Agreements template operations

        [Function(nameof(DownloadAgreementsTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_agreements_template", tags: new[] { "events" }, Summary = "Download the agreements template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "agreementId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the agreement.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadAgreementsTemplate(
                        [HttpTrigger(AuthorizationLevel.Anonymous, "GET",
                            Route = "bodies/{bodyId:required}/events/{eventId:required}/agreements/{agreementId:required}/downloadTemplate")] HtttestquestData req,
                        string bodyId, string eventId, string agreementId)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var agreementTemplate = await _eventAgreementService.DownloadAgreementsTemplate(ctx, bodyId, eventId, agreementId, true);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"); // .docx
            await response.WriteBytesAsync(agreementTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        #endregion

        #region AgendaItems template operations

        [Function(nameof(DownloadAgendaItemsTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_agendaitems_template", tags: new[] { "events" }, Summary = "Download the agenda items template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadAgendaItemsTemplate(
                       [HttpTrigger(AuthorizationLevel.Anonymous, "GET",
                            Route = "bodies/{bodyId:required}/events/{eventId:required}/agendaItems/downloadTemplate")] HtttestquestData req,
                       string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            var agendaItemsTemplate = await _eventAgendaItemsService.DownloadAgendaItemsTemplate(ctx, bodyId, eventId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/pdf"); // .pdf
            response.Headers.Add("Content-Disposition", "attachment; filename=filename.pdf"); // Uncomment to download in swagger
            await response.WriteBytesAsync(agendaItemsTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion

        #region Votation Service

        [Function(nameof(GetEventVotationDetail))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_votationdetail", tags: new[] { "events" }, Summary = "Return the details of an event votation.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventVotation[]), Summary = "")]
        public async Task<HtttestsponseData> GetEventVotationDetail([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/votationDetail")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var votationDetails = await this._eventVotationService.GetAllVotationsDetailByEvent(ctx, bodyId, eventId, fromArchive, true);

            var response = votationDetails != null ? req.CreateResponse(HttpStatusCode.OK) : req.CreateResponse(HttpStatusCode.Forbidden);
            await response.WriteStringAsync(JsonConvert.SerializeObject(votationDetails)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(AddOrUpdateEventVotationDetail))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "addorupdate_event_votationdetail", tags: new[] { "events" }, Summary = "Add or update the votation details for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(EventVotation[]), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventVotation[]), Summary = "")]
        public async Task<HtttestsponseData> AddOrUpdateEventVotationDetail([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/votationDetail")] HtttestquestData req, string bodyId, string eventId)
        {
            var votationDetailToUpdate = JsonConvert.DeserializeObject<EventVotation[]>(await new StreamReader(req.Body).ReadToEndAsync()) ?? throw new ArgumentException("Invalid request body");
            using var ctx = await CreatePnPContextAsUser(req);

            var agendaItems = await this._eventVotationService.AddOrUpdateVotationForEvent(ctx, bodyId, eventId, votationDetailToUpdate);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(agendaItems)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetEventVotingRetestsentation))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_votationretestsentation", tags: new[] { "events" }, Summary = "Return the voting retestsentation for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Dictionary<string, IEnumerable<string>>), Summary = "")]
        public async Task<HtttestsponseData> GetEventVotingRetestsentation([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/votingRetestsentation")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(req);
            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            var votingRetestsentation = await this._eventVotationService.GetRetestsentationsByUser(ctx, bodyId, eventId, fromArchive);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(votingRetestsentation)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(StartEventVotation))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_event_votationretestsentation", tags: new[] { "events" }, Summary = "Start the votation process for an event.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(IEnumerable<EventVotation>), Summary = "")]
        public async Task<HtttestsponseData> StartEventVotation([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "bodies/{bodyId:required}/events/{eventId:required}/startVotation")] HtttestquestData req, string bodyId, string eventId)
        {
            using var ctx = await CreatePnPContextAsUser(req);
            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);

            var votations = await this._eventVotationService.StartVotation(ctx, bodyId, eventId, true);

            var response = votations != null ? req.CreateResponse(HttpStatusCode.OK) : req.CreateResponse(HttpStatusCode.Forbidden);
            await response.WriteStringAsync(JsonConvert.SerializeObject(votations)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateVoteByUser))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_event_voteByUser", tags: new[] { "events" }, Summary = "Update the vote submitted by a user.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(Vote), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(IEnumerable<EventVotation>), Summary = "")]

        public async Task<HtttestsponseData> UpdateVoteByUser([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/updateVote")] HtttestquestData req, string bodyId, string eventId)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser(req);
                var vote = JsonConvert.DeserializeObject<Vote>(await new StreamReader(req.Body).ReadToEndAsync()) ?? throw new ArgumentException("Invalid request body");

                var votations = await this._eventVotationService.UpdateVoteByUser(ctx, bodyId, eventId, vote.Id, vote.VoteId, userUpn);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(votations)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateVoteByRetestsentation))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_event_voteByRetestsentation", tags: new[] { "events" }, Summary = "Update the vote submitted by a retestsentation.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "retestsentationId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the retestsentation.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(Vote), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(EventVotation), Summary = "")]

        public async Task<HtttestsponseData> UpdateVoteByRetestsentation([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "bodies/{bodyId:required}/events/{eventId:required}/updateVoteByRetestsentation/{retestsentationId:required}")] HtttestquestData req, string bodyId, string eventId, string retestsentationId)
        {
            using var ctx = await CreatePnPContextAsUser(req);
            var vote = JsonConvert.DeserializeObject<Vote>(await new StreamReader(req.Body).ReadToEndAsync()) ?? throw new ArgumentException("Invalid request body");

            var votations = await this._eventVotationService.UpdateVoteByRetestsentation(ctx, bodyId, eventId, vote.Id, vote.VoteId, retestsentationId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(votations)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }




        #endregion

        #region Sensitivity Labels Service
        [Function(nameof(GetTenantSensitivityLabels))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_tenant_sensitivity_labels", tags: new[] { "sensitivitylabels" }, Summary = "Return all sensitivity labels from the current tenant.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TenantSensitivityLabel[]), Summary = "")]
        public async Task<HtttestsponseData> GetTenantSensitivityLabels([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "sensitivitylabels")] HtttestquestData req)
        {
            var result = await _eventFilesSensitivityLabelsService.GetTenantSensitivityLabelsAsync();
            var response = result != null || result!.Any() ? req.CreateResponse(HttpStatusCode.OK) : req.CreateResponse(HttpStatusCode.Forbidden);
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetDocumentsSensitivityLabelsByEventSharedId))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_documents_sensitivity_labels", tags: new[] { "sensitivitylabels" }, Summary = "Return all documents from an event with its respective sensitivity labels.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiParameter(name: "isEditor", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the user is editor or not.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TenantSensitivityLabel[]), Summary = "")]
        public async Task<HtttestsponseData> GetDocumentsSensitivityLabelsByEventSharedId([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId:required}/events/{eventId:required}/documentSensitivityLabels")] HtttestquestData req, string bodyId, string eventId)
        {
            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            _ = bool.TryParse(req.Query["isEditor"], out var isEditor);
            var result = await _eventFilesSensitivityLabelsService.GetDocumentsSensitivityLabelsByEventSharedId(bodyId, eventId, fromArchive, isEditor);
            var response = result != null ? req.CreateResponse(HttpStatusCode.OK) : req.CreateResponse(HttpStatusCode.Forbidden);
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateSensitivityLabel))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_sensitivity_label", tags: new[] { "sensitivitylabels" }, Summary = "Return all sensitivity labels from the current tenant.")]
        [OpenApiParameter(name: "fromArchive", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the event is from the archived library.")]
        [OpenApiParameter(name: "isEditor", In = ParameterLocation.Query, Required = false, Type = typeof(bool), Summary = "Indicates if the user is editor or not.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(TenantSensitivityLabel[]), Summary = "")]
        public async Task<HtttestsponseData> UpdateSensitivityLabel([HttpTrigger(AuthorizationLevel.Anonymous, "PATCH", Route = "sensitivitylabels")] HtttestquestData req)
        {
            _ = bool.TryParse(req.Query["fromArchive"], out var fromArchive);
            _ = bool.TryParse(req.Query["isEditor"], out var isEditor);
            var documentSensitivityLabel = JsonConvert.DeserializeObject<DocumentSensitivityLabel>(await new StreamReader(req.Body).ReadToEndAsync());
            var result = await _eventFilesSensitivityLabelsService.UpdateSensitivityLabel(documentSensitivityLabel!, fromArchive, isEditor);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion
    }

    #region Event publishing orchestration

    public static class OrchestrationHelpers
    {
        public static string GetInstanceIdForEventChanges(string bodyId, string eventId) => $"es-b:{bodyId.Trim().ToLowerInvariant()}-e:{eventId.Trim().ToLowerInvariant()}";

        public static int? GetCaptureIdFromInstanceId(string instanceId)
        {
            if (int.TryParse(instanceId.Substring(instanceId.LastIndexOf("-c:") + 3), out var enumCapture))
            {
                return enumCapture;
            }
            return null;
        }

        public static async Task<bool> IsPendingOrchestrations(DurableTaskClient client, string bodyId, string eventId, Func<OrchestrationMetadata, bool>? condition = null)
        {
            var existingInstances = client.GetAllInstancesAsync(new OrchestrationQuery()
            {
                InstanceIdprefix = GetInstanceIdForEventChanges(bodyId, eventId),
                Statuses = [OrchestrationRuntimeStatus.Running, OrchestrationRuntimeStatus.Pending, OrchestrationRuntimeStatus.Suspended],
                PageSize = 10
            });

            var enumerator = existingInstances.GetAsyncEnumerator();
            var hasElement = await enumerator.MoveNextAsync() && enumerator.Current != null;
            if (!hasElement)
                return false;

            if (condition == null)
                return true;

            do
            {
                if (condition(enumerator.Current!))
                    return true;
                else
                    continue;
            } while (await enumerator.MoveNextAsync() && enumerator.Current != null);

            return false;
        }
    }

    public class DurableEventSendingOrchestration : TriggerBase<DurableEventSendingOrchestration>
    {
        private EventPublishingService _eventPublishingService;

        public DurableEventSendingOrchestration(EventPublishingService eventPublishingService, ILogger<DurableEventSendingOrchestration> logger, M365AuthHelper auth) : base(logger, auth)
        {
            this._eventPublishingService = eventPublishingService;
        }

        [Function(nameof(DurableEventSendingOrchestration))]
        public async Task RunOrchestrator([OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var logger = context.CreateReplaySafeLogger<DurableEventSendingOrchestration>();
            var operation = context.GetInput<EventSendOperationInformation>()!;

            using var scope = logger.BeginContosoScope("Event publishing orchestration", operation.BodyId, operation.SharedEventId, operation.SentBy);
            var startTime = context.CurrentUtcDateTime;

            try
            {
                while (await context.CallShouldWaitFortestviosprocessAsync(operation))
                {
                    logger.LogInformation("Waiting for testvious process to finish to continue with: " + context.InstanceId);
                    await context.CreateTimer(context.CurrentUtcDateTime.AddMinutes(1), CancellationToken.None);
                }

                // ensure event group and outlook meeting exists  
                await context.CallprocessprepareRequirementsActivityAsync(operation);

                await context.CallprocessDocumentSetChangesActivityAsync(operation, new TaskOptions() { Retry = new TaskRetryOptions(new RetryPolicy(3, TimeSpan.FromSeconds(5))) });

                var outlookChanges = context.CallprocessOutlookUpdatessActivityAsync(operation);
                var attendanceChanges = context.CallprocessTasksUpdatesActivityAsync(operation);
                var notificationsChanges = context.CallprocessNotificationsActivityAsync(operation);

                logger.LogDebug("Waiting for outlook, attendance and notifications jobs");
                await Task.WhenAll(outlookChanges, attendanceChanges, notificationsChanges);
                logger.LogDebug("Waiting for outlook, attendance and notifications jobs finished");

                await context.CallprocessRemovalOrArchiveActivityAsync(operation);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, $"Error processing: {ex.Message}");
            }
            finally
            {
                logger.LogInformation($"Finished event publishing orchestration for {context.InstanceId} in {context.CurrentUtcDateTime.Subtract(startTime).TotalMinutes} min");
            }
        }

        [Function(nameof(ShouldWaitFortestviosprocess))]
        public async Task<bool> ShouldWaitFortestviosprocess([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext, [DurableClient] DurableTaskClient client)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (ShouldWaitFortestviosprocess)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            var result = false;

            logger.LogDebug($"Starting processing ShouldWaitFortestviosprocess");
            try
            {
                var currentCaptureId = operation.ToCaptureState!.CaptureId;

                result = await OrchestrationHelpers.IsPendingOrchestrations(client, operation.BodyId, operation.SharedEventId, (metadata) =>
                {
                    var enumCapture = OrchestrationHelpers.GetCaptureIdFromInstanceId(metadata.InstanceId);
                    var isTimeout = metadata.LastUpdatedAt.AddMinutes(30) < DateTime.UtcNow; // orchestration takes 5 min max

                    if (enumCapture.HasValue && currentCaptureId > enumCapture && !isTimeout)
                    {
                        logger.LogInformation($"Found existing pending job: {metadata.InstanceId}");
                        return true;
                    }

                    return false;
                });
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing ShouldWaitFortestviosprocesss");
            }
            finally
            {
                logger.LogDebug($"Exiting processing ShouldWaitFortestviosprocess");
            }

            return result;
        }

        [Function(nameof(processprepareRequirementsActivity))]
        public async Task processprepareRequirementsActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processprepareRequirementsActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing prepareRequirementsActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.prepareRequirements(ctx, operation);

                logger.LogInformation($"Finished processing prepareRequirementsActivity");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing prepare requirements");
            }
            finally
            {
                logger.LogDebug($"Exiting processing prepareRequirementsActivity");
            }
        }

        [Function(nameof(processDocumentSetChangesActivity))]
        public async Task processDocumentSetChangesActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processDocumentSetChangesActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing DocumentSetChangesActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.EventCaptureToDocumentSetUpdates(ctx, operation);

                logger.LogInformation($"Finished processing DocumentSetChangesActivity");
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                logger.LogDebug($"Exiting processing DocumentSetChangesActivity");
            }
        }

        [Function(nameof(processOutlookUpdatessActivity))]
        public async Task processOutlookUpdatessActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processOutlookUpdatessActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing OutlookUpdatessActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.EventCaptureToTeamsAndOutlookUpdates(ctx, operation);

                logger.LogInformation($"Finished processing OutlookUpdatessActivity");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing outlook");
            }
            finally
            {
                logger.LogDebug($"Exiting processing OutlookUpdatessActivity");
            }
        }

        [Function(nameof(processTasksUpdatesActivity))]
        public async Task processTasksUpdatesActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processTasksUpdatesActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing TasksUpdatesActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.EventCaptureToTasksUpdates(ctx, operation);

                logger.LogInformation($"Finished processing TasksUpdatesActivity");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing tasks");
            }
            finally
            {
                logger.LogDebug($"Exiting processing TasksUpdatesActivity");
            }
        }

        [Function(nameof(processNotificationsActivity))]
        public async Task processNotificationsActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processNotificationsActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing processNotificationsActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.EventCaptureToNotifications(ctx, operation);

                logger.LogInformation($"Finished processing processNotificationsActivity");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing notifications");
            }
            finally
            {
                logger.LogDebug($"Exiting processing processNotificationsActivity");
            }
        }

        [Function(nameof(processRemovalOrArchiveActivity))]
        public async Task processRemovalOrArchiveActivity([ActivityTrigger] EventSendOperationInformation operation, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableEventSendingOrchestration>();
            using var scope = logger.BeginContosoScope("Event publishing orchestration (processRemovalOrArchiveActivity)", operation.BodyId, operation.SharedEventId, operation.SentBy);

            logger.LogDebug($"Starting processing processRemovalOrArchiveActivity");
            try
            {
                using var ctx = await CreatePnPContextAsSystem() as PnPContext;

                await _eventPublishingService.EventCaptureToRemovalOrArchive(ctx, operation);

                logger.LogInformation($"Finished processing processRemovalOrArchiveActivity");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error processing removal or archive");
            }
            finally
            {
                logger.LogDebug($"Exiting processing processRemovalOrArchiveActivity");
            }
        }
    }

    #endregion
}