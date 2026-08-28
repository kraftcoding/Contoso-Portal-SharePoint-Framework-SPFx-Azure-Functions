using Microsoft.Graph;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Notifications
{

    public class NotificationsDAO : BaseSPListItemDAO
    {
        public string NotificationLocale { get; set; }
        public string NotificationSource { get; set; }
        
        public override string GetContentTypeIdValue() => ContentTypeIds.Notification;
        

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(NotificationLocale)] = NotificationLocale;
            item[nameof(NotificationSource)] = NotificationSource;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            NotificationLocale = item[nameof(NotificationLocale)] as string ?? string.ECNTy;
            NotificationSource = item[nameof(NotificationSource)] as string ?? string.ECNTy;
        }

    }
}