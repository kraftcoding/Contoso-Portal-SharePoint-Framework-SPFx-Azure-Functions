using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Management;
using Contoso.Portal.Model.profile;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.profileHttpTrigger
{
    public class ManagementHttpTriger : TriggerBase<ManagementHttpTriger>
    {
        private ManagementService _managementService;
        private ConfigDepartmentsService _ConfigDepartmentsService;
        public ManagementHttpTriger(ManagementService managementService,
            ConfigDepartmentsService ConfigDepartmentsService,
            ILogger<ManagementHttpTriger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            _managementService = managementService;
            _ConfigDepartmentsService = ConfigDepartmentsService;
        }

        #region Cache
        [Function(nameof(CleanBodyCache))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "clean_body_cache", tags: new[] { "users" }, Summary = "")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> CleanBodyCache([HttpTrigger(AuthorizationLevel.Anonymous, "DELETE", Route = "management/{bodyId:required}/cache/{bodyToClean:required}")] HtttestquestData req, string bodyId, string bodyToClean)
        {

            var userUpn = GetUpnFromRequest(req);

            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
                await this._managementService.CleanBodyCache(ctx, userUpn, bodyId, bodyToClean);
                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(string.ECNTy)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion

        #region EXTERNAL Bodies
        [Function(nameof(GetEXTERNALBodyConfig))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get-EXTERNAL-body-config", tags: new[] { "bodies" }, Summary = "")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(ConfigEXTERNALBodies[]), Summary = "")]
        public async Task<HtttestsponseData> GetEXTERNALBodyConfig([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId}/EXTERNALBodies")] HtttestquestData req, string bodyId)
        {

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

                var ConfigDepartments = await this._managementService.GetEXTERNALBodiesInformation(ctx, userUpn, bodyId);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(ConfigDepartments)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateEXTERNALBodyInfo))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_EXTERNAL_body_info", tags: new[] { "bodies" }, Summary = "Update EXTERNAL body information")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> UpdateEXTERNALBodyInfo([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId}/EXTERNALBody")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var bodyInfo = JsonConvert.DeserializeObject<ConfigEXTERNALBodies>(await new StreamReader(req.Body).ReadToEndAsync()); ;

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at UpdateEXTERNALBodyInfo");

            if (bodyInfo is null)
                throw new ArgumentNullException("Invalid request body at UpdateEXTERNALBodyInfo");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            await _managementService.UpdateEXTERNALBodyInformation(ctx, userUpn, bodyId, bodyInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(CreateEXTERNALBody))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "create_EXTERNAL_body_info", tags: new[] { "bodies" }, Summary = "Create EXTERNAL body information")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> CreateEXTERNALBody([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "management/{bodyId}/EXTERNALBody")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var bodyInfo = JsonConvert.DeserializeObject<ConfigEXTERNALBodies>(await new StreamReader(req.Body).ReadToEndAsync()); ;

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at CreateEXTERNALBody");

            if (bodyInfo is null)
                throw new ArgumentNullException("Invalid request body at CreateEXTERNALBody");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var result = await _managementService.CreateEXTERNALBody(ctx, userUpn, bodyId, bodyInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(DownloadEXTERNALCertificateTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_certificate_template", tags: new[] { "management" }, Summary = "Download the attendance template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the current body.")]
        [OpenApiParameter(name: "certificateBodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body from which the certificate is to be downloaded.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadEXTERNALCertificateTemplate([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId:required}/EXTERNALBody/downloadTemplate/{certificateBodyId:required}")] HtttestquestData req, string bodyId, string certificateBodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at DownloadCertificateTemplate");

            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);
            var certificateTemplate = await _managementService.DownloadEXTERNALCertificateTemplate(ctx, userUpn, bodyId, certificateBodyId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/pdf"); // .pdf
            response.Headers.Add("Content-Disposition", "attachment; filename=filename.pdf"); // Uncomment to download in swagger
            await response.WriteBytesAsync(certificateTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion


        #region Bodies

        [Function(nameof(GetBodyConfig))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "GetBodyConfig", tags: new[] { "bodies" }, Summary = "")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(ConfigDepartments[]), Summary = "")]
        public async Task<HtttestsponseData> GetBodyConfig([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId}/bodies")] HtttestquestData req, string bodyId)
        {

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

                var ConfigDepartments = await this._managementService.GetBodiesInformation(ctx, userUpn, bodyId);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(ConfigDepartments)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(CheckPermissions))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "check-permissions", tags: new[] { "bodies" }, Summary = "")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(ConfigDepartments[]), Summary = "")]
        public async Task<HtttestsponseData> CheckPermissions([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "management/permissions")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var bodyRoleInfo = JsonConvert.DeserializeObject<UserBodyRole>(await new StreamReader(req.Body).ReadToEndAsync());

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at UpdateBodyUsers");

            if (bodyRoleInfo is null)
                throw new ArgumentNullException("Invalid request body at UpdateBodyUsers");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyRoleInfo.BodyId}", req);
            PermissionCheckResult result = await _managementService.CheckPermissions(ctx, bodyRoleInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/json");
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateBodyInfo))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_body_info", tags: new[] { "bodies" }, Summary = "Update body information")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> UpdateBodyInfo([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId}/body")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var bodyInfo = JsonConvert.DeserializeObject<ConfigDepartments>(await new StreamReader(req.Body).ReadToEndAsync()); ;

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at UpdateBodyInfo");

            if (bodyInfo is null)
                throw new ArgumentNullException("Invalid request body at UpdateBodyInfo");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            await _managementService.UpdateBodyInformation(ctx, userUpn, bodyId, bodyInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateBodyUsers))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_body_users", tags: new[] { "bodies" }, Summary = "Update body information")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> UpdateBodyUsers([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId}/body/users")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var bodyRoleInfo = JsonConvert.DeserializeObject<List<UserBodyRole>>(await new StreamReader(req.Body).ReadToEndAsync());

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at UpdateBodyUsers");

            if (bodyRoleInfo is null)
                throw new ArgumentNullException("Invalid request body at UpdateBodyUsers");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var (added, error) = await _managementService.AddBodyUsersRolesRequest(ctx, userUpn, bodyId, bodyRoleInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync($"Added: {string.Join(", ", added)}{(error.Count > 0 ? string.Join(", ", error) : "")}").ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        // [Function(nameof(DownloadCertificateTemplate))]
        // [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        // [OpenApiOperation(operationId: "download_certificate_template", tags: new[] { "management" }, Summary = "Download the attendance template.")]
        // [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the current body.")]
        // [OpenApiParameter(name: "certificateBodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the body from which the certificate is to be downloaded.")]
        // [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        // public async Task<HtttestsponseData> DownloadCertificateTemplate([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId:required}/body/downloadTemplate/{certificateBodyId:required}")] HtttestquestData req, string bodyId, string certificateBodyId)
        // {
        //     var userUpn = GetUpnFromRequest(req);
        //     if (string.IsNullOrECNTy(userUpn))
        //         throw new ArgumentException("ECNTy UPN at DownloadCertificateTemplate");

        //     using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);
        //     var certificateTemplate = await _managementService.DownloadCertificateTemplate(ctx, userUpn, bodyId, certificateBodyId);

        //     var response = req.CreateResponse(HttpStatusCode.OK);
        //     response.Headers.Add("Content-Type", "application/pdf"); // .pdf
        //     response.Headers.Add("Content-Disposition", "attachment; filename=filename.pdf"); // Uncomment to download in swagger
        //     await response.WriteBytesAsync(certificateTemplate).ConfigureAwait(false);
        //     return await Task.FromResult(response).ConfigureAwait(false);
        // }
        #endregion

        #region Users

        [Function(nameof(CreateUser))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_user_info", tags: new[] { "users" }, Summary = "Create a user on Entra ID")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> CreateUser([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "management/{bodyId:required}/user")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var newUser = JsonConvert.DeserializeObject<BodyUserInfo>(await new StreamReader(req.Body).ReadToEndAsync()); ;

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at CreateUser");

            if (newUser == null)
                throw new ArgumentNullException("Invalid request body at CreateUser");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var result = await this._managementService.CreateUser(ctx, userUpn, bodyId, newUser);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(GetUsersInfo))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "GetUsersInfo", tags: new[] { "users" }, Summary = "")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(Dictionary<string, UserInfo>), Summary = "")]
        public async Task<HtttestsponseData> GetUsersInfo([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId}/users")] HtttestquestData req, string bodyId)
        {

            var userUpn = GetUpnFromRequest(req);

            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

                var members = await this._managementService.GetUsersInformation(ctx, userUpn, bodyId);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(members)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(UpdateUserInfo))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "update_user_info", tags: new[] { "users" }, Summary = "Update user information on EntraID")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> UpdateUserInfo([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId}/user")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var userInfo = JsonConvert.DeserializeObject<UserInfo>(await new StreamReader(req.Body).ReadToEndAsync()); ;

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at UpdateUserInfo");

            if (userInfo == null)
                throw new ArgumentNullException("Invalid request body at UpdateUserInfo");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            await this._managementService.UpdateUserInformation(ctx, userUpn, bodyId, userInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(string.ECNTy).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }


        [Function(nameof(GetInvitationText))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_invitation_text", tags: new[] { "Invitations" }, Summary = "Get default invitation text")]
        [OpenApiParameter(name: "serverRelativeUrl", In = ParameterLocation.Query, Required = true, Type = typeof(string), Summary = "Server relative URL of the file")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "text/plain", bodyType: typeof(string), Summary = "Returns the invitation text")]
        public async Task<HtttestsponseData> GetInvitationText(
            [HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/invitationtext")] HtttestquestData req)
        {
            var query = System.Web.HttpUtility.ParseQueryString(req.Url.Query);
            var serverRelativeUrl = query["serverRelativeUrl"];

            var response = req.CreateResponse();

            if (string.IsNullOrECNTy(serverRelativeUrl))
            {
                response.StatusCode = HttpStatusCode.BadRequest;
                await response.WriteStringAsync("Missing parameter: serverRelativeUrl");
                return response;
            }

            try
            {
                var text = await this._managementService.GetInvitationText(serverRelativeUrl);
                response.StatusCode = HttpStatusCode.OK;
                await response.WriteStringAsync(text);
            }
            catch (Exception ex)
            {
                response.StatusCode = HttpStatusCode.InternalServerError;
                await response.WriteStringAsync($"Error: {ex.Message}");
            }

            return response;
        }


        #endregion

        #region Licenses
        [Function(nameof(GetAvailableLicenses))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_available_licenses", tags: new[] { "LicenseInfo" }, Summary = "Get the licenses available at the tenant")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(LicenseInfo[]), Summary = "")]
        public async Task<HtttestsponseData> GetAvailableLicenses([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId}/availableLicenses")] HtttestquestData req, string bodyId)
        {

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn) || string.IsNullOrECNTy(bodyId))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
                var result = await _managementService.GetAvailableLicenses(ctx, userUpn, bodyId);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(ManageUserLicense))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "manage_user_license", tags: new[] { "UserInfo" }, Summary = "Update licenses for a user")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> ManageUserLicense([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId:required}/manageUserLicense")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var userInfo = JsonConvert.DeserializeObject<UserInfo>(await new StreamReader(req.Body).ReadToEndAsync());

            if (string.IsNullOrECNTy(userUpn) || string.IsNullOrECNTy(bodyId))
                throw new ArgumentException("Missing parameters at ManageUserLicense");

            if (userInfo is null)
                throw new ArgumentNullException("Invalid request body at ManageUserLicense");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var result = await _managementService.ManageUserLicense(ctx, userUpn, bodyId, userInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(result)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion

        #region BusinessUnit
        [Function(nameof(GetBusinessUnit))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "get_sectorial_area_ministry", tags: new[] { "BusinessAreaMinistries" }, Summary = "Get update sectorial area ministry")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(BusinessUnit[]), Summary = "")]
        public async Task<HtttestsponseData> GetBusinessUnit([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "management/{bodyId}/BusinessUnit")] HtttestquestData req, string bodyId)
        {

            var userUpn = GetUpnFromRequest(req);
            HtttestsponseData response;
            if (string.IsNullOrECNTy(userUpn))
            {
                response = req.CreateResponse(HttpStatusCode.BadRequest);
            }
            else
            {
                using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);

                var BusinessAreaMinistries = await this._managementService.GetAllBusinessAreasMinistries(ctx, userUpn, bodyId);

                response = req.CreateResponse(HttpStatusCode.OK);
                await response.WriteStringAsync(JsonConvert.SerializeObject(BusinessAreaMinistries)).ConfigureAwait(false);
            }

            return await Task.FromResult(response).ConfigureAwait(false);
        }

        [Function(nameof(AddOrUpdateBusinessUnit))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "add_or_update_sectorial_area_ministry", tags: new[] { "BusinessAreaMinistries" }, Summary = "Add or update sectorial area ministry")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "Id of the current body")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> AddOrUpdateBusinessUnit([HttpTrigger(AuthorizationLevel.Anonymous, "PUT", Route = "management/{bodyId}/BusinessUnit")] HtttestquestData req, string bodyId)
        {
            var userUpn = GetUpnFromRequest(req);
            var BusinessUnitInfo = JsonConvert.DeserializeObject<BusinessUnit>(await new StreamReader(req.Body).ReadToEndAsync());

            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at AddOrUpdateBusinessUnit");

            if (BusinessUnitInfo is null)
                throw new ArgumentNullException("Invalid request body at AddOrUpdateBusinessUnit");

            using var ctx = await CreatePnPContextAsUser($"/sites/{bodyId}", req);
            var BusinessUnit = await _managementService.AddOrUpdateBusinessUnit(ctx, userUpn, bodyId, BusinessUnitInfo);
            var response = req.CreateResponse(HttpStatusCode.OK);
            await response.WriteStringAsync(JsonConvert.SerializeObject(BusinessUnit)).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }



        #endregion
    }
}