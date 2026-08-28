using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksApprodvalMinutesDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALproviderForJsonFiles<TasksApprodvalMinutesDAO, TasksApprodvalMinutesJsonDAO, TasksApprodvalMinutesDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<TasksApprodvalMinutesDTO> GetTaskById(IPnPContext ctx, int taskId)
    {
        return await GetDTOById(ctx, taskId);
    }

    public async Task<IEnumerable<TasksApprodvalMinutesDTO>> AddOrUpdateForEvent(IPnPContext ctx, string sharedEventId, IEnumerable<TasksApprodvalMinutesDTO> taksDTOs)
    {
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, sharedEventId);
        return await AddOrUpdateInFolder(ctx, folder, taksDTOs);
    }

    public async Task<IEnumerable<TasksApprodvalMinutesDTO>> GetByEventId(IPnPContext ctx, int eventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathById(ctx, eventId);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);
        return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);
    }

    public async Task<IEnumerable<TasksApprodvalMinutesDTO>> GetPendingTaskBySharedEventId(IPnPContext bodyCtx, string sharedEventId)
    {
        var wssid = await bodyCtx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery("EstadoRequest", "Integer", wssid.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task<IEnumerable<TasksApprodvalMinutesDTO>> GetAllTasksBySharedEventId(IPnPContext bodyCtx, string sharedEventId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    public async Task DeleteBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        var approdvals = await GetByView(ctx, DefaultView, hiddenFolderServerRelativePath);

        foreach (var app in approdvals)
            await app.AsListItem().DeleteAsync();
    }
}