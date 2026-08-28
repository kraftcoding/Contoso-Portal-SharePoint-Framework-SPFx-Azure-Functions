using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksCertificationDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALproviderForJsonFiles<TasksCertificationDAO, TasksCertificationJsonDAO, TasksCertificationDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<TasksCertificationDTO> GetTaskById(IPnPContext ctx, int taskId)
    {
        return await GetDTOById(ctx, taskId);
    }

    public async Task<IEnumerable<TasksCertificationDTO>> AddOrUpdateForEvent(IPnPContext ctx, string shareEventId, IEnumerable<TasksCertificationDTO> taksDTOs)
    {
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, shareEventId);
        return await AddOrUpdateInFolder(ctx, folder, taksDTOs);
    }

    public async Task<IEnumerable<TasksCertificationDTO>> GetByEventId(IPnPContext ctx, int eventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathById(ctx, eventId);

        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);
    }

    public async Task<IEnumerable<TasksCertificationDTO>> GetPendingByEvent(IPnPContext ctx, string sharedEventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        // TODO: check remove in operator?
        var wssidTermPending = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var estadoRequest = new List<int>() { wssidTermPending };

        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksCertificationDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);
    }

    public async Task<IEnumerable<TasksCertificationDTO>> GetAllTasksBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);
        var wssidTermPending = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var wssidTermCancelled = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Cancelled);
        var wssidTermRejected = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Rejected);
        var wssidTermPendingDelegation = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.PendingDelegation);
        var wssidTermDelegated = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Delegated);
        var wssidTermAccepted = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Accepted);
        var wssidTermAcceptedBySystem = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.AcceptedBySystem);
        var wssidTermExpired = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Expired);
        var wssidTermPendingModification = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.PendingModification);
        var estadoRequest = new List<int>() { wssidTermPending, wssidTermCancelled, wssidTermRejected, wssidTermPendingDelegation, wssidTermDelegated, wssidTermAccepted, wssidTermAcceptedBySystem, wssidTermExpired, wssidTermPendingModification };

        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksCertificationDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);
    }
}