using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public class EventAgreementDocumentDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.DocumentCertificacion;

        public override void MapToSPListItem(IListItem item)
        {

            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
        }
    }
}