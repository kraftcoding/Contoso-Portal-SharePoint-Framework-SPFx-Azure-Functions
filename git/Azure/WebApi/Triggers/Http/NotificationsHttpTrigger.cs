using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Model.Notifications;
using Contoso.Portal.Runtime;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.NotificationsHttpTrigger
{
    public class NotificationsHttpTrigger : TriggerBase<NotificationsHttpTrigger>
    {
        private NotificationsService _notificationsService;
        private Settings Settings; // refactor

        public NotificationsHttpTrigger(NotificationsService notificationsService,
            Settings settings,
            ILogger<NotificationsHttpTrigger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            this._notificationsService = notificationsService;
            this.Settings = settings;
        }

        [Function(nameof(GetUserNotifications))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_notifications", tags: new[] { "notifications" }, Summary = "Return a collection of unread notifications for the current user.")]
        [OpenApiParameter(name: "locale", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Language to filter the user notifications.")]
        [OpenApiParameter(name: "source", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Source to filter the user notifications.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserNotifications([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "notifications")] HtttestquestData req, string? locale, string? source)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrWhiteSpace(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                var requestLocale = this.Settings.DefaultLocale;
                using var ctx = await CreatePnPContextAsSystem();
                var notifications = await this._notificationsService.GetUserNotifications(ctx, userUpn, requestLocale!, source);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(notifications)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(GetUserArchivedNotifications))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_notifications_archived", tags: new[] { "notifications" }, Summary = "Return a collection of archived notifications for the current user.")]
        [OpenApiParameter(name: "locale", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Language to filter the user notifications.")]
        [OpenApiParameter(name: "source", In = ParameterLocation.Query, Required = false, Type = typeof(string), Summary = "Source to filter the user notifications.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserArchivedNotifications([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "notificationsArchived")] HtttestquestData req, string? locale, string? source)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrWhiteSpace(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                var requestLocale = this.Settings.DefaultLocale;
                using var ctx = await CreatePnPContextAsSystem();
                var notifications = await this._notificationsService.GetUserArchivedNotifications(ctx, userUpn, requestLocale!, source);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(notifications)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetUserNotificationById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_notification", tags: new[] { "notifications" }, Summary = "Return a specific notification by ID.")]
        [OpenApiParameter(name: "notificationId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the notification.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification), Summary = "")]
        public async Task<HtttestsponseData> GetUserNotificationById([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "notifications/{notificationId:required}")] HtttestquestData req, string notificationId)
        {
            var userUpn = GetUpnFromRequest(req);
            Guid notificationGuid;
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn) || !Guid.TryParse(notificationId, out notificationGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                var notification = await this._notificationsService.GetNotificationById(ctx, userUpn, notificationGuid);
                if (notification == null)
                {
                    response = req.CreateResponse(HttpStatusCode.NotFound);
                }
                else
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                    await response.WriteStringAsync(JsonConvert.SerializeObject(notification)).ConfigureAwait(false);
                }
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateVisualizedUserNotificationById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_notification", tags: new[] { "notifications" }, Summary = "Update the read status of the notification for the current user.")]
        [OpenApiParameter(name: "notificationId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the notification.")]
        [OpenApiParameter(name: "readed", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Indicates if the notification has been read.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification), Summary = "")]
        public async Task<HtttestsponseData> UpdateVisualizedUserNotificationById([HttpTrigger(AuthorizationLevel.Anonymous, "PATCH", Route = "notifications/{notificationId:required}/setStatus/{readed:required}")] HtttestquestData req, string notificationId, bool readed)
        {
            var userUpn = GetUpnFromRequest(req);
            Guid notificationGuid;
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn) || !Guid.TryParse(notificationId, out notificationGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                var notification = await this._notificationsService.UpdateVisualizedUserNotificationById(ctx, userUpn, notificationGuid, readed);

                if (notification == null)
                {
                    response = req.CreateResponse(HttpStatusCode.NotFound);
                }
                else
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                }
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(HideUserNotificationById))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_notification", tags: new[] { "notifications" }, Summary = "Hide a specific notification by ID for the current user.")]
        [OpenApiParameter(name: "notificationId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the notification.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification), Summary = "")]
        public async Task<HtttestsponseData> HideUserNotificationById([HttpTrigger(AuthorizationLevel.Anonymous, "PATCH", Route = "notifications/{notificationId:required}/hide")] HtttestquestData req, string notificationId)
        {
            var userUpn = GetUpnFromRequest(req);
            Guid notificationGuid;
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn) || !Guid.TryParse(notificationId, out notificationGuid))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                var notification = await this._notificationsService.HideNotificationById(ctx, userUpn, notificationGuid);

                if (notification == null)
                {
                    response = req.CreateResponse(HttpStatusCode.NotFound);
                }
                else
                {
                    response = req.CreateResponse(HttpStatusCode.OK);
                }
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateNotificationInBatch))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "batch_notification", tags: new[] { "notifications" }, Summary = "Updates user notifications in batch.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(NotificationBatchRequest), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification), Summary = "")]
        public async Task<HtttestsponseData> UpdateNotificationInBatch([HttpTrigger(AuthorizationLevel.Anonymous, "PATCH", Route = "notifications/batch")] HtttestquestData req)
        {
            var request = JsonConvert.DeserializeObject<NotificationBatchRequest>(await new StreamReader(req.Body).ReadToEndAsync());

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;

            if (string.IsNullOrECNTy(userUpn) || request == null || !request.IsValidRequest())
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                await _notificationsService.UpdateNotifications(ctx, userUpn, request);
                response = req.CreateResponse(HttpStatusCode.OK);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(AddUrgentNotification))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "AddUrgentNotification", tags: new[] { "notifications" }, Summary = "Create a new urgent notification.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "eventId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the event.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(Notification), Required = true)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Notification), Summary = "")]
        public async Task<HtttestsponseData> AddUrgentNotification([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "notifications/bodies/{bodyId:required}/event/{eventId:required}/UrgentNotification")] HtttestquestData req, string bodyId, string eventId)
        {
            var notification = JsonConvert.DeserializeObject<Notification>(await new StreamReader(req.Body).ReadToEndAsync());
            var userUpn = GetUpnFromRequest(req);

            HtttestsponseData response;
            if (string.IsNullOrWhiteSpace(userUpn) || string.IsNullOrWhiteSpace(eventId) || notification == null)
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var rootCtx = await CreatePnPContextAsSystem();

                await this._notificationsService.UrgentNotification(rootCtx, bodyId, notification, userUpn, eventId);

                response = req.CreateResponse(HttpStatusCode.OK);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }
    }
}