using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Model.Events;

namespace Contoso.Portal.Data.DTO.Notifications
{
    public class NotificationDTO : JsonFileListItemAbstractionDTO<NotificationsDAO, NotificationJsonDAO>
    {
        public string? Title { get => JsonDAO.Title; set => JsonDAO.Title = value; }
        public DateTime Expires { get => JsonDAO.Expires; set => JsonDAO.Expires = value; }
        public Guid Id { get => JsonDAO.Id; set => JsonDAO.Id = value; }
        public string? NotificationType { get => JsonDAO.NotificationType; set => JsonDAO.NotificationType = value; }
        public string? NotificationPriority { get => JsonDAO.NotificationPriority; set => JsonDAO.NotificationPriority = value; }
        public string? Body { get => JsonDAO.Body; set => JsonDAO.Body = value; }

        public List<PendingChanges>? PendingChanges { get => JsonDAO.PendingChanges; set => JsonDAO.PendingChanges = value; }
        public string? Source
        {
            get
            {
                return ItemDAO.NotificationSource;
            }

            set
            {
                JsonDAO.Source = value;
                ItemDAO.NotificationSource = value;
            }
        }

        public string? Locale
        {
            get
            {
                return ItemDAO.NotificationLocale;
            }

            set
            {
                JsonDAO.Locale = value;
                ItemDAO.NotificationLocale = value;
            }
        }
        public string? Url { get => JsonDAO.Url; set => JsonDAO.Url = value; }
        public bool Readed { get => JsonDAO.Readed; set => JsonDAO.Readed = value; }
        public DateTime Created { get => ItemDAO.Created; }
    }
}
