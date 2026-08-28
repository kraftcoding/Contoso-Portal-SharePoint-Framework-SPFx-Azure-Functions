using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.DemoEntities;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;


namespace Contoso.Portal.Triggers.Http.DemoEntitiesHttpTrigger
{
    public class DemoEntitiesHttpTrigger : TriggerBase<DemoEntitiesHttpTrigger>
    {
        private DemoEntitiesService _service;
        public DemoEntitiesHttpTrigger(DemoEntitiesService service,
            ILogger<DemoEntitiesHttpTrigger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            this._service = service;
        }


        #region Registro de Entidades de Memoria Democrática
        [Function(nameof(DownloadDemoEntityCertificateTemplate))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "download_remd_certificate_template", tags: new[] { "DemoEntities" }, Summary = "Download the attendance template.")]
        [OpenApiParameter(name: "bodyId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the current body.")]
        [OpenApiParameter(name: "entidadId", In = ParameterLocation.Path, Required = true, Type = typeof(string), Summary = "ID of the entity from which the certificate is to be downloaded.")]
        [OpenApiResponseWithoutBody(statusCode: HttpStatusCode.OK, Summary = "")]
        public async Task<HtttestsponseData> DownloadDemoEntityCertificateTemplate([HttpTrigger(AuthorizationLevel.Anonymous, "GET", Route = "DemoEntities/{bodyId:required}/downloadTemplate/{entidadId:required}")] HtttestquestData req, string bodyId, string entidadId)
        {
            var userUpn = GetUpnFromRequest(req);
            if (string.IsNullOrECNTy(userUpn))
                throw new ArgumentException("ECNTy UPN at DownloadCertificateTemplate");

            using var ctx = await CreatePnPContextAsUser(BodyPatternUtilities.BodyIdToSiteUrl(bodyId), req);

            var certificateTemplate = await _service.DownloadDemoEntityCertificateTemplate(ctx, userUpn, bodyId, entidadId);

            var response = req.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Type", "application/pdf"); // .pdf
            response.Headers.Add("Content-Disposition", "attachment; filename=filename.pdf"); // Uncomment to download in swagger
            await response.WriteBytesAsync(certificateTemplate).ConfigureAwait(false);
            return await Task.FromResult(response).ConfigureAwait(false);
        }
        #endregion

    }

}