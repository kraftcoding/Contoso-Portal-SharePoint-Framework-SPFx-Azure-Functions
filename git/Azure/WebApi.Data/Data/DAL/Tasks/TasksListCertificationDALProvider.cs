using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksListCertificationDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALprovider<TasksCertificationDAO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);


    public async Task<IEnumerable<TasksCertificationDAO>> GetByEventId(IPnPContext ctx, int eventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathById(ctx, eventId);

        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetByView(ctx, view, hiddenFolderServerRelativePath);
    }

    public async Task<IEnumerable<TasksCertificationDAO>> GetPendingByEvent(IPnPContext ctx, string sharedEventId)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);

        // TODO: check remove in operator?
        var wssidTermPending = await ctx.Web.GetWssIdForTermAsync(TaxonomyValuesIds.RequestStatus.Pending);
        var estadoRequest = new List<int>() { wssidTermPending };

        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(TasksCertificationDTO.EstadoRequest), "Integer", estadoRequest)}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);

        return await GetByView(ctx, view, hiddenFolderServerRelativePath);
    }
}