using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Event
{
    public class EventPublishedDocumentDALprovider : BaseSPListDALprovider<EventPublishedDocumentDAO>
    {
        public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.PublishedEvents;

        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

        public async Task<IEnumerable<EventPublishedDocumentDAO>> GetPublishedDocument(IPnPContext ctx, string sharedEventId)
        {
            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(BaseEventDAO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
            return await GetByView(ctx, view);
        }

    }

    public class EventInConstructionDocumentDALprovider : BaseSPListDALprovider<EventInConstructionDocumentDAO>
    {
        public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

        public async Task<IEnumerable<EventInConstructionDocumentDAO>> GetInConstructionDocument(IPnPContext ctx, string sharedEventId)
        {
            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(BaseEventDAO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
            return await GetByView(ctx, view);
        }

    }

}