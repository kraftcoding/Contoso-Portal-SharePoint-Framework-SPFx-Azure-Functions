using DocumentFormat.OpenXml.Stestadsheet;
using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAL.Management;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Runtime.Helpers;
using Contoso.Portal.Triggers.Http.EventsHttpTrigger;
using Contoso.Portal.Triggers.Http.PowerBiReportsHttpTrigger;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Timer;

public class PowerBiReportsTrigger(
   PowerBiReportsService reportsService,
   ILogger<PowerBiReportsTrigger> logger,
   M365AuthHelper auth
   ) : TriggerBase<PowerBiReportsTrigger>(logger, auth)
{
    private readonly PowerBiReportsService _reportsService = reportsService;

    [Function(nameof(PowerBiReportsMonitoringJob))]
    public async Task PowerBiReportsMonitoringJob([TimerTrigger("0 0 0 * * *")] FunctionContext context, [DurableClient] DurableTaskClient client)
    {
        using var rootCtx = await CreatePnPContextAsSystem();
        var dal = new processReportsDALprovider();
        var result = await dal.GetLastSuccesfulExecution(rootCtx);
        await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableReportExecutionOrchestration), result.Inicioprodceso.HasValue ? result.Inicioprodceso.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ") : null);
    }

}