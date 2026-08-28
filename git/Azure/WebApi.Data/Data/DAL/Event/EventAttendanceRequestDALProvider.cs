using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DTO.Event;

namespace Contoso.Portal.Data.DAL.Event;

public class EventAttendanceRequestDALprovider(EventDALprovider eventsDALprovider) : BaseDALproviderForEventJsonFiles<EventAttendanceDAO, EventAttendanceJsonDAO, EventAttendanceDTO>(eventsDALprovider)
{
    public override string ListSiteRelavteUrl => EventDALprovider.ListSiteRelavteUrl;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
}