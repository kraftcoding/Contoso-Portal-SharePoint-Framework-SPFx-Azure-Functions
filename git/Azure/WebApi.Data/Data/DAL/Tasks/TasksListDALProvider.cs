using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksListDALprovider() : BaseSPListDALprovider<TasksDAO>()
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } =
    [
        (nameof(BaseSPListItemDAO.Created), false)
    ];

    public async Task<IEnumerable<TasksDAO>> GetPendingTask(IPnPContext ctx)
    {
        var estadoRequest = new List<int>();
        var wssidTermPending = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        estadoRequest.Add(wssidTermPending);
        //var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><Or><Or><Or>{CAML_GetOperatorQuery(nameof(TasksModificationMinutesDTO.ContentTypeId), "ContentTypeId", ContentTypeIds.Task.RequestAttendance, ComparisonOperators.BeginsWith)}{CAML_GetOperatorQuery(nameof(TasksModificationMinutesDTO.ContentTypeId), "ContentTypeId", ContentTypeIds.Task.RequestDelegation, ComparisonOperators.BeginsWith)}</Or>{CAML_GetOperatorQuery(nameof(TasksModificationMinutesDTO.ContentTypeId), "ContentTypeId", ContentTypeIds.Task.RequestCertificacion, ComparisonOperators.BeginsWith)}</Or><Or>{CAML_GetOperatorQuery(nameof(TasksModificationMinutesDTO.ContentTypeId), "ContentTypeId", ContentTypeIds.Task.RequestModificacion, ComparisonOperators.BeginsWith)}{CAML_GetOperatorQuery(nameof(TasksModificationMinutesDTO.ContentTypeId), "ContentTypeId", ContentTypeIds.Task.RequestAprodbacionMinutes, ComparisonOperators.BeginsWith)}</Or></Or>{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var taskDAO = await GetByView(ctx, view);

        return taskDAO;
    }


}