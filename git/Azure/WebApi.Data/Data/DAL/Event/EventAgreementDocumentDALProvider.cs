using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Event
{
    public class EventAgreementDocumentDALprovider : BaseSPListDALprovider<EventAgreementDocumentDAO>
    {
        private readonly EventDALprovider _eventsDALprovider;

        public EventAgreementDocumentDALprovider(EventDALprovider eventsDALprovider)
        {
            this._eventsDALprovider = eventsDALprovider;
        }

        public override string ListSiteRelavteUrl => _eventsDALprovider.ListSiteRelavteUrl;
        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

        public async Task SetContentTypeAgreementDocument(IPnPContext ctx, string uIdDocument)
        {
            var item = await GetByUniqueId(ctx, uIdDocument) ?? throw new Exception($"Error in {nameof(SetContentTypeAgreementDocument)} {nameof(uIdDocument)}:'{uIdDocument}' not found");
            item.ContentTypeId = await GetContentTypeIdValueInList(ctx);
            await item.AsListItem().UpdateAsync();
        }
    }
}