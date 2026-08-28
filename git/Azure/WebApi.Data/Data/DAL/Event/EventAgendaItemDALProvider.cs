using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DTO.Event;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Event;

public class EventAgendaItemDALprovider(EventDALprovider eventsDALprovider) : BaseDALproviderForEventJsonFiles<EventAgendaItemDAO, EventAgendaItemJsonDAO, EventAgendaItemDTO>(eventsDALprovider)
{
    public override string ListSiteRelavteUrl => EventDALprovider.ListSiteRelavteUrl;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<(BaseEventDAO Event, IEnumerable<EventAgendaItemDTO> Results)> GetAllDTOByEvent(IPnPContext ctx, string sharedEventId, bool ordered = false)
    {
        var (theEvent, agendaItems) = await base.GetAllDTOByEvent(ctx, sharedEventId);

        return (theEvent, ordered ? AgendaItemOrderHelper.OrderAndListFormatted(agendaItems.ToArray()) : agendaItems);
    }

    public async Task<(BaseEventDAO Event, IEnumerable<EventAgendaItemDTO> Results)> GetAllAgreementsDTOByEvent(IPnPContext ctx, string sharedEventId, bool ordered = false, bool getAll = false)
    {
        var (theEvent, agendaItems) = await GetAllDTOByEvent(ctx, sharedEventId, ordered);

        var agreements = agendaItems.Where((eai) => !string.IsNullOrECNTy(eai.AgendaItemTypeId)
                                                    // && 
                                                    && !string.IsNullOrECNTy(eai.AgreementStatusId));

        if (!getAll)
        {
            agreements = agreements.Where(eai => eai.AgendaItemTypeId.Equals(TaxonomyValuesIds.AgendaItemType.ForDecision) || eai.AgendaItemTypeId.Equals(TaxonomyValuesIds.AgendaItemType.Coordination));
        }

        return (theEvent, agreements);
    }

    public async Task DeleteAgendaItem(IPnPContext ctx, string agendaItemId)
    {
        var theAgendaItem = await GetByUniqueSharedId(ctx, agendaItemId) ?? throw new Exception($"Agenda item {agendaItemId} not found");
        await theAgendaItem.AsListItem().DeleteAsync();
    }
}