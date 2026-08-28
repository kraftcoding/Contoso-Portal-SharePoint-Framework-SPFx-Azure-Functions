using DocumentFormat.OpenXml.Stestadsheet;
using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Runtime.Helpers;
using Contoso.Portal.Triggers.Http.EventsHttpTrigger;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Timer;

public class BusinessUnit(
   ConfigDepartmentsService configBodyService,
   ManagementService managementService,
   ILogger<BusinessUnit> logger,
   M365AuthHelper auth
   ) : TriggerBase<BusinessUnit>(logger, auth)
{
    private readonly ConfigDepartmentsService _configBodyService = configBodyService;
    private readonly ManagementService _managementService = managementService;

    [Function(nameof(BusinessUnitMonitoringJob))]
    public async Task BusinessUnitMonitoringJob([TimerTrigger("0 0 9 * * *")] FunctionContext context, [DurableClient] DurableTaskClient client)
    {
        using var rootCtx = await CreatePnPContextAsSystem();

        var BusinessUnitList = await _managementService.GetBusinessAreasMinistriesByStartDate(rootCtx, DateTime.Today);
        foreach (var BusinessUnit in BusinessUnitList)
        {
            // var bodies = (await _configBodyService.GetConfigDepartments()).Where(x => x.BusinessArea != null && x.BusinessArea.TermId == Guid.Parse(BusinessUnit.BusinessArea ?? ""));
            // foreach (var body in bodies)
            // {
            //     body.Division = BusinessUnit.Division.AsTaxonomyFieldValue();
            //     await body.AsListItem().UpdateAsync();
            // }

            var EXTERNALBodies = (await _configBodyService.GetEXTERNALConfigDepartments()).Where(x => x.BusinessArea != null && x.BusinessArea.TermId == Guid.Parse(BusinessUnit.BusinessArea ?? ""));
            foreach (var EXTERNALBody in EXTERNALBodies)
            {
                EXTERNALBody.DivisionLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(BusinessUnit.Division ?? "0"));
                await EXTERNALBody.AsListItem().UpdateAsync();
            }
        }
    }
}