using Google.prodtobuf.WellKnownTypes;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Attributes;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Enums;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Domains.profile;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Management;
using Contoso.Portal.Model.profile;
using Contoso.Portal.Runtime.Helpers;
using Newtonsoft.Json;
using System.Net;

namespace Contoso.Portal.Triggers.Http.PowerBiReportsHttpTrigger
{
    public class PowerBiReportsHttpTrigger : TriggerBase<PowerBiReportsHttpTrigger>
    {
        private PowerBiReportsService _reportsService;
        public PowerBiReportsHttpTrigger(PowerBiReportsService reportsService,
            ILogger<PowerBiReportsHttpTrigger> logger,
            M365AuthHelper auth)
            : base(logger, auth)
        {
            _reportsService = reportsService;
        }

        #region Reports
        [Function(nameof(RunReport))]
        [OpenApiSecurity("bearer_auth", SecuritySchemeType.Http, Scheme = OpenApiSecuritySchemeType.Bearer, BearerFormat = "JWT")]
        [OpenApiOperation(operationId: "run_report", tags: new[] { "powerbireports" }, Summary = "Launch a new instance to dump data for PowerBi reports")]
        [OpenApiParameter(name: "startDate", In = ParameterLocation.Path, Required = false, Type = typeof(string), Summary = "")]
        [OpenApiResponseWithBody(statusCode: HttpStatusCode.OK, contentType: "application/json", bodyType: typeof(string), Summary = "")]
        public async Task<HtttestsponseData> RunReport([HttpTrigger(AuthorizationLevel.Anonymous, "POST", Route = "powerbireports/{startDate?}")] HtttestquestData req, string? startDate, [DurableClient] DurableTaskClient client)
        {

            var userUpn = GetUpnFromRequest(req);
            var instanceId = $"re-d:{DateTime.UtcNow.ToTimestamp()}-rt:{(string.IsNullOrECNTy(startDate) ? DALConstants.Management.TypeOfReport.Full.ToLower() : DALConstants.Management.TypeOfReport.Incremental.ToLower())}";

            await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableReportExecutionOrchestration), startDate);

            var response = req.CreateResponse(HttpStatusCode.Accepted);
            await response.WriteStringAsync($"Started orchestration with ID = '{instanceId}'.");
            return response;
        }
        #endregion

    }

    public class DurableReportExecutionOrchestration : TriggerBase<DurableReportExecutionOrchestration>
    {
        private readonly PowerBiReportsService _reportService;

        public DurableReportExecutionOrchestration(PowerBiReportsService reportService, ILogger<DurableReportExecutionOrchestration> logger, M365AuthHelper auth)
            : base(logger, auth)
        {
            _reportService = reportService;
        }

        [Function(nameof(DurableReportExecutionOrchestration))]
        public async Task RunOrchestrator([OrchestrationTrigger] TaskOrchestrationContext context)
        {
            var logger = context.CreateReplaySafeLogger<DurableReportExecutionOrchestration>();
            var startDate = context.GetInput<string>();


            using var scope = logger.BeginContosoScope("Report execution orchestration");
            var startTime = context.CurrentUtcDateTime;

            try
            {
                logger.LogInformation($"Starting durable orchestration for PowerBi report");

                // Llamada a la actividad que ejecuta el informe
                var result = await context.CallActivityAsync<string>(nameof(processReportExecutionActivity), startDate);

                logger.LogInformation($"Report execution result: {result}");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error running PowerBi report execution orchestration");
            }
            finally
            {
                logger.LogInformation($"Finished report orchestration {context.InstanceId} in {context.CurrentUtcDateTime.Subtract(startTime).TotalMinutes} minutes");
            }
        }

        [Function(nameof(processReportExecutionActivity))]
        public async Task processReportExecutionActivity([ActivityTrigger] string startDate, FunctionContext executionContext)
        {
            var logger = executionContext.GetLogger<DurableReportExecutionOrchestration>();
            using var scope = logger.BeginContosoScope("PowerBi Report execution activity");

            try
            {
                logger.LogInformation("Running report.");
                using var rootCtx = await CreatePnPContextAsSystem();
                await _reportService.RunReport(rootCtx, startDate);
                logger.LogInformation("Report executed successfully.");
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error executing report");
                throw;
            }
        }
    }

}