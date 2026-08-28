namespace Contoso.Portal.Model.Notifications
{
    public class NotificationUserConfig
    {
        public int Id { get; set; }
        public Guid unreadUserNotificationsUniqueId { get; set; }
        public Guid hiddenUserNotificationsUniqueId { get; set; }
        private string unreadUserFolderName
        {
            get { return $"{unreadUserNotificationsUniqueId:N}"; }
        }
        private string readedUserFolderName
        {
            get { return $"{hiddenUserNotificationsUniqueId:N}"; }
        }

        public string GetPathFromStatus(bool status)
        {
            return status ? readedUserFolderName : unreadUserFolderName;
        }
    }
}
