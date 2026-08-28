using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.BodiesHttpTrigger
{
    public class BodiesHttpTrigger : TriggerBase<BodiesHttpTrigger>
    {
        private BodiesService _bodiesService;
        private ManagementService _bodyManagementService;
        public BodiesHttpTrigger(BodiesService bodiesService, ManagementService bodyManagementService, ILogger<BodiesHttpTrigger> logger, M365AuthHelper auth)
            : base(logger, auth)
        {
            this._bodiesService = bodiesService;
            _bodyManagementService = bodyManagementService;
        }

        [Function(nameof(GetUserBodies))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_bodies", tags: new[] { "bodies" }, Summary = "Return all bodies of the current user.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(LocalizedUserBodyRoles[]), Summary = "")]
        public async Task<HtttestsponseData> GetUserBodies([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/my")] HtttestquestData req)
        {
            using var ctx = await CreatePnPContextAsUser(req);

            var bodies = await this._bodiesService.GetUserInitialBodiesInformation(ctx);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(bodies)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetBodyUsersByRole))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_body_usersByRole", tags: new[] { "bodies" }, Summary = "Return all users of a body by role.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body.")]
        [OpenApiParameter(name: "role", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "Role of the users to filter by.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(User[]), Summary = "")]
        public async Task<HtttestsponseData> GetBodyUsersByRole([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "bodies/{bodyId}/users")] HtttestquestData req, string bodyId)
        {
            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var roleString = req.Query["role"];
            var role = BodyPatternUtilities.RoleFromString(roleString ?? throw new ArgumentNullException("Role shoud not be null"));
            if (role == BodyRole.Undefined)
                throw new ArgumentException($"Role {roleString} is not implemented");

            var members = await this._bodiesService.GetBodyUsersByRole(ctx, bodyId, role);

            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(members)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        // Añadir más endpoints relacionados con bodies, para la gestin de correos


    }
}