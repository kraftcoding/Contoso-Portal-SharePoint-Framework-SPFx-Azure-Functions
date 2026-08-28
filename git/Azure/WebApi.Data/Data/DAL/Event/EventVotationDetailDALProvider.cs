using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DTO.Event;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Event;

public class EventVotationDetailDALprovider(EventDALprovider eventsDALprovider)
    : BaseDALproviderForEventJsonFiles<EventVotationDetailDAO, EventVotationJsonDAO, EventVotationDetailDTO>(eventsDALprovider)
{
    public override string ListSiteRelavteUrl => EventDALprovider.ListSiteRelavteUrl;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<IEnumerable<EventVotationDetailDTO>> GetVotationDetailForEventSharedIdAndIdUser(IPnPContext bodyCtx, string sharedEventId, int lookupId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(EventVotationDetailDAO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery(nameof(EventVotationDetailDAO.AsignadoA), "Integer", lookupId.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var task = await GetDTOByView(bodyCtx, view);
        return task;
    }

    [Obsolete("mover a base listitem; es el mismo metodo para cualqueir item de lista")]
    public async Task DeleteVotationDetails(IPnPContext ctx, int voteSessionId)
    {
        var votationDetailsDTO = await GetById(ctx, voteSessionId) ?? throw new Exception($"Votation id {voteSessionId} not found");
        await votationDetailsDTO.AsListItem().DeleteAsync();
    }
}