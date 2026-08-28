using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Notifications
{
    public class NotificationMessagesDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.NotificationConfig;
        public Boolean Activo { get; set; }

    }
}