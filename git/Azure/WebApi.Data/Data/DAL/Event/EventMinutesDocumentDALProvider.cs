using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Event
{
    public class EventMinutesDocumentDALprovider : BaseSPListDALprovider<EventMinutesDocumentDAO>
    {
        private readonly EventDALprovider _eventsDALprovider;

        public EventMinutesDocumentDALprovider(EventDALprovider eventsDALprovider)
        {
            this._eventsDALprovider = eventsDALprovider;
        }

        public override string ListSiteRelavteUrl => _eventsDALprovider.ListSiteRelavteUrl;
        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

        public async Task SetContentTypeMinutesDocument(IPnPContext ctx, string uIdDocument)
        {
            var item = await GetByUniqueId(ctx, uIdDocument) ?? throw new Exception($"Error in {nameof(SetContentTypeMinutesDocument)} {nameof(uIdDocument)}:'{uIdDocument}' not found");
            item.ContentTypeId = await GetContentTypeIdValueInList(ctx);
            await item.AsListItem().UpdateAsync();
        }
    }
}