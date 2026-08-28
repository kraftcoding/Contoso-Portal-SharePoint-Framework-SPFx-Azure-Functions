using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksModificationMinutesDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALproviderForJsonFiles<TasksModificationMinutesDAO, TasksModificationMinutesJsonDAO, TasksModificationMinutesDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<TasksModificationMinutesDTO> GetTaskById(IPnPContext ctx, int taskId)
    {
        return await GetDTOById(ctx, taskId);
    }

    public async Task<IEnumerable<TasksModificationMinutesDTO>> GetAllTasksBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        // var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        // var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        // return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);

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

        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetDTOByView(ctx, view, hiddenFolderServerRelativePath);
    }

    public async Task AddorUpdateForEvent(IPnPContext ctx, string sharedEventId, TasksModificationMinutesDTO taskDTO)
    {
        //1. Get Folder
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, sharedEventId);

        //2. Save File taskDelegate
        var listItem = new List<TasksModificationMinutesDTO> { taskDTO };
        await AddOrUpdateInFolder(ctx, folder, listItem);
    }

    public async Task DeletePendingTasksByEvent(IPnPContext ctx, string sharedEventId)
    {
        //1. Get Folder
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        //2. Get PendingTask Value
        // TODO: check remove in operator?
        var wssidTermPending = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var estadoRequest = new List<int> { wssidTermPending };
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksModificationMinutesDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort);

        //3. Delete PendingTask
        var tasks = await GetByView(ctx, view, hiddenFolderServerRelativePath);

        //TODO: Paralelizar            
        foreach (var att in tasks)
            await att.AsListItem().DeleteAsync();
    }
}