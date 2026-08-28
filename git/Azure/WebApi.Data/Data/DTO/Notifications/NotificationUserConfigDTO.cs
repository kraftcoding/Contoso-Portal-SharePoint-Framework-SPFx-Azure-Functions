using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Notifications;

namespace Contoso.Portal.Data.DTO.Notifications
{
    public class NotificationUserConfigDTO : JsonFileListItemAbstractionDTO<NotificationUserConfigDAO, NotificationUserConfigJsonDAO>
    {
        public string? unreadUserNotificationsUniqueId { get => JsonDAO.unreadUserNotificationsUniqueId; set => JsonDAO.unreadUserNotificationsUniqueId = value; }
        public string? readedUserNotificationsUniqueId { get => JsonDAO.readedUserNotificationsUniqueId; set => JsonDAO.readedUserNotificationsUniqueId = value; }
    }
}
