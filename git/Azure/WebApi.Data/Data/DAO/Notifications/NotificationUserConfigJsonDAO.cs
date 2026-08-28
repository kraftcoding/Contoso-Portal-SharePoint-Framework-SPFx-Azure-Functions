
using Newtonsoft.Json;

namespace Contoso.Portal.Data.DAO.Notifications
{
    public class NotificationUserConfigJsonDAO : BaseJsonDataDAO
    {
        [JsonIgnore]
        public string userId { get; set; }
        public string? unreadUserNotificationsUniqueId { get; set; }
        public string? readedUserNotificationsUniqueId { get; set; }

        public override string GetFileName() => $"nuc-{userId}.json";
        public override bool IsECNTy() => false;
    }    
}
