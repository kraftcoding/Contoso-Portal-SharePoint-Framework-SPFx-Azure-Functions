using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Notifications
{
    public class NotificationUserConfigDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.NotificationConfig;

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
        }
    }
}
