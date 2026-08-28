using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksDALprovider(Source source = Source.InConstruction) : BaseSPListDALproviderForJsonFiles<TasksDAO, TasksJsonDAO, TasksDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = new();

    private readonly Source _source = source;

    public override string ListSiteRelavteUrl => _source switch
    {
        Source.InConstruction => ListsSiteRelativeUrls.InConstructionEvents,
        Source.Archived => ListsSiteRelativeUrls.ArchivedEvents,
        Source.Published => throw new NotImplementedException(),
        _ => throw new NotImplementedException()
    };

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        (nameof(BaseSPListItemDAO.Created), false)
    };

    public async Task<IEnumerable<TasksDTO>> GetFilterTask(IPnPContext ctx, string bodyId)
    {
        using var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(bodyId)));
        var estadoRequest = new List<int>();
        var wssidTermPending = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        estadoRequest.Add(wssidTermPending);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var taskDTO = await GetDTOByView(bodyCtx, view);
        return taskDTO;
    }

    //TODO: no usar haasta prodximo Alert. 
    public async Task<IEnumerable<TasksDTO>> GetFilterPendingTask(IPnPContext ctx, string bodyId)
    {
        using var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(bodyId)));
        var estadoRequest = new List<int>();
        var wssidTermPending = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        estadoRequest.Add(wssidTermPending);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var taskDTO = await GetDTOByView(bodyCtx, view);
        return taskDTO;
    }

    public async Task<IEnumerable<TasksDTO>> GetFilterTaskByEvent(IPnPContext bodyCtx, string sharedEventId)
    {
        var wssid = await bodyCtx.Web.GetWssIdForTermAsync(DALConstants.TaxonomyValuesIds.RequestStatus.Cancelled);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery("EstadoRequest", "Integer", wssid.ToString(), ComparisonOperators.Neq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task<IEnumerable<TasksDTO>> GetTaskByIdMeetingAndIdUser(IPnPContext bodyCtx, string sharedEventId, int lookupId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery(nameof(TasksDTO.AsignadoA), "Integer", lookupId.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }
    //TODO: Revisar si existe alguna tarea delegada al usuario a borrar. Si es asi marcar la tarea delegada como pendiente.

    public async Task<IEnumerable<TasksDTO>> GetDelegatedTaskByIdMeetingAndIdUserDelegated(IPnPContext bodyCtx, string sharedEventId, int lookupId)
    {
        var wssid = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Delegated);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery(nameof(TasksDTO.Delegado), "Integer", lookupId.ToString(), ComparisonOperators.Eq, true)}</And>{CAML_GetOperatorQuery("EstadoRequest", "Integer", wssid.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task<IEnumerable<TasksDTO>> GetCancelTaskBySharedEventIdAndIdUser(IPnPContext bodyCtx, string sharedEventId, int lookupId)
    {
        var wssid = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Cancelled);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery(nameof(TasksDTO.AsignadoA), "Integer", lookupId.ToString(), ComparisonOperators.Eq, true)}</And>{CAML_GetOperatorQuery("EstadoRequest", "Integer", wssid.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task<IEnumerable<TasksDTO>> GetPendingTaskByEvent(IPnPContext bodyCtx, string sharedEventId)
    {
        var wssid = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery("EstadoRequest", "Integer", wssid.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task<IEnumerable<TasksDTO>> AddOrUpdateForEvent(IPnPContext ctx, string sharedEventId, IEnumerable<TasksDTO> tasks)
    {
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, sharedEventId);

        return await AddOrUpdateInFolder(ctx, folder, tasks);
    }
}