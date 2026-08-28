using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DTO.Event;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Event;

public class EventMinutesInformationDALprovider(EventDALprovider eventsDALprovider) : BaseDALproviderForEventJsonFiles<EventMinutesInformationDAO, EventMinutesInformationJsonDAO, EventMinutesInformationDTO>(eventsDALprovider)
{
    public override string ListSiteRelavteUrl => EventDALprovider.ListSiteRelavteUrl;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<EventMinutesInformationDTO?> AddOrUpdateForEvent(IPnPContext ctx, string sharedEventId, EventMinutesInformationDTO minutesInfoDTO)
    {
        return (await AddOrUpdateItemsForEvent(ctx, sharedEventId, [minutesInfoDTO])).FirstOrDefault();
    }

    public async Task<EventMinutesInformationDTO?> GetBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var (_, results) = await GetAllDTOByEvent(ctx, sharedEventId);
        return results.FirstOrDefault();
    }
}