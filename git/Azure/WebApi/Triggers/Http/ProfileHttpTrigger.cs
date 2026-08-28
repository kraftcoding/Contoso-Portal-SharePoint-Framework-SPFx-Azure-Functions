using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.profile;
using Contoso.Portal.Model.profile;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.profileHttpTrigger
{
    public class profileHttpTrigger : TriggerBase<profileHttpTrigger>
    {
        private profileService _profileService;
        public profileHttpTrigger(profileService profileService,
            ILogger<profileHttpTrigger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            this._profileService = profileService;
        }

        [Function(nameof(UpdateMemberprofile))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "Update_profile", tags: new[] { "profile" }, Summary = "Update the profile information of a user.")]
        [OpenApiRequestBody(contentType: "application/json", bodyType: typeof(profileInfo), Required = true)]
        public async Task<HtttestsponseData> UpdateMemberprofile([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "profile")] HtttestquestData req)
        {
            var profile = JsonConvert.DeserializeObject<profileInfo>(await new StreamReader(req.Body).ReadToEndAsync());
            if (profile == null)
                throw new ArgumentException("Invalid request body");

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                await this._profileService.Updateprofile(ctx, userUpn, profile);
                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            }
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetMemberprofile))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "Get_profile", tags: new[] { "profile" }, Summary = "Return the profile information of a user.")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(profileInfo), Summary = "")]
        public async Task<HtttestsponseData> GetMemberprofile([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "profile")] HtttestquestData req)
        {
            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsSystem();
                var updateMember = await this._profileService.Getprofile(ctx, userUpn);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(updateMember)).ConfigureAwait(false);
            }
            return await Task.FromResult(response).ConfigureAwait(false);
        }
    }
}