using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksDelegateDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALproviderForJsonFiles<TasksDelegateDAO, TasksDelegateJsonDAO, TasksDelegateDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<TasksDelegateDTO> GetTaskById(IPnPContext ctx, int taskId)
    {
        return await GetDTOById(ctx, taskId);
    }

    public async Task AddorUpdateForEvent(IPnPContext ctx, string sharedEventId, TasksDelegateDTO taskDTOs)
    {
        //1. Get Folder
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, sharedEventId);

        //2. Save File taskDelegate
        await AddOrUpdateInFolder(ctx, folder, [taskDTOs]);
    }

    public async Task<IEnumerable<TasksDelegateDTO>> GetAllTasksBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        //TODO: Revisar por qué si se añaden los estados, la consulta funciona. Si se eliminan, no funciona.
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