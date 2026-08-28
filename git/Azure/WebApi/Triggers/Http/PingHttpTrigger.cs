using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Domains.Tasks;

using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.PingHttpTrigger
{
    public class PingHttpTrigger : TriggerBase<PingHttpTrigger>
    {
        private readonly IServiceProvider _serviceProvider;

        public PingHttpTrigger(ILogger<PingHttpTrigger> logger,
            M365AuthHelper auth,
            IServiceProvider serviceProvider)
            : base(logger, auth)
        {
            this._serviceProvider = serviceProvider;
        }

        [Function(nameof(PingAuthenticated))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "pingAuthenticated", tags: new[] { "diagnostics" }, Summary = "Simple health prodbe with authentication.", Description = "Simple health prodbe with authentication.", Visibility = OpenApiVisibilityType.Important)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(String), Summary = "Will return user Microsoft Identity header.", Description = "Will return user Microsoft Identity header.")]
        public async Task<HtttestsponseData> PingAuthenticated([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "ping")] HtttestquestData req)
        {
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync((string)JsonConvert.SerializeObject(req.Headers)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(PingAnonymous))]
        [OpenApiOperation(operationId: "pingAnonymous", tags: new[] { "diagnostics" }, Summary = "Simple health prodbe.", Description = "Simple health prodbe.", Visibility = OpenApiVisibilityType.Important)]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(String), Summary = "Will return ping.", Description = "Will return ping.")]
        public async Task<HtttestsponseData> PingAnonymous([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "noauth/ping")] HtttestquestData req)
        {
            using var scope = _serviceProvider.CreateScope();
            var services = scope.ServiceProvider;

            // TODO: Add your services here
            /*
                        var bodyRoleService = services.GetRequiredService<BodyRoleService>();
                        var bodiesService = services.GetRequiredService<BodiesService>();
                        var ConfigDepartmentsService = services.GetRequiredService<ConfigDepartmentsService>();
                        var eventsService = services.GetRequiredService<EventsService>();
                        var eventMinutesService = services.GetRequiredService<EventMinutesService>();
                        var eventMinutesTemplateService = services.GetRequiredService<EventMinutesTemplateService>();
                        var eventAttendanceService = services.GetRequiredService<EventAttendanceService>();
                        var eventCaptureService = services.GetRequiredService<EventCaptureService>();
                        var eventAgreementService = services.GetRequiredService<EventAgreementService>();
                        var ProfileService = services.GetRequiredService<ProfileService>();
                        var tasksService = services.GetRequiredService<TasksService>();
                        var tasksAttendanceService = services.GetRequiredService<TasksAttendanceService>();
                        var tasksDelegateService = services.GetRequiredService<TasksDelegateService>();
                        var tasksApprodvalMinutesService = services.GetRequiredService<TasksApprovalMinutesService>();
                        var tasksModificationMinutesService = services.GetRequiredService<TasksModificationMinutesService>();
                        var tasksCertificationService = services.GetRequiredService<TasksCertificationService>();
                        var notificationsBusinessService = services.GetRequiredService<NotificationsBusinessService>();
                        var notificationsService = services.GetRequiredService<NotificationsService>();
                        var m365AuthHelper = services.GetRequiredService<M365AuthHelper>();
                        var taxonomyTranslationService = services.GetRequiredService<TaxonomyTranslationService>();
                        var eventPublishingService = services.GetRequiredService<EventPublishingService>();
                        var simService = services.GetRequiredService<INotificationSmsService>();
            */

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync((string)JsonConvert.SerializeObject(req.Headers)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
    }
}