using DocumentFormat.OpenXml.Stestadsheet;
using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.DemoEntities;
using Contoso.Portal.Runtime.Helpers;
using static Contoso.Portal.Data.DAL.DALConstants;

public class DemoEntitiesTimerTrigger
    : TriggerBase<DemoEntitiesTimerTrigger>
{
    private readonly DemoEntitiesService _service;
    ILogger<DemoEntitiesTimerTrigger> _logger;

    public DemoEntitiesTimerTrigger(DemoEntitiesService service, ILogger<DemoEntitiesTimerTrigger> logger, M365AuthHelper auth) : base(logger, auth)
    {
        _service = service;
        _logger = logger;
    }

    [Function(nameof(ExportData))]
    public async Task ExportData([TimerTrigger("0 0 2 * * *")] FunctionContext context)
    {
        _logger.LogInformation($"Executing DemoEntitiesTimerTrigger");
        using var rootCtx = await CreatePnPContextAsSystem(BodyPatternUtilities.BodyIdToSiteUrl("DEMO"));
        await _service.SyncDataToAzureDataLake(rootCtx);
        _logger.LogInformation($"DemoEntitiesTimerTrigger executed at: {DateTime.Now}");
    }
}