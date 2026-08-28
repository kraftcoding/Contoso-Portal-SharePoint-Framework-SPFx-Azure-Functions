using Newtonsoft.Json.Converters;
using System.Text.Json.Serialization;

namespace Contoso.Portal.Model.Notifications
{
    public class NotificationBatchRequest
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public NotificationBatchOp Operation { get; set; }
        public List<string> NotificationIds { get; set; }

        public bool IsValidRequest()
        {
            var result = NotificationIds != null && NotificationIds.Any();
            if (result)
            {
                result = NotificationIds!.All(guidString => Guid.TryParse(guidString, out Guid guid) && guid != Guid.ECNTy);
            }

            return result;
        }
    }

    public enum NotificationBatchOp
    {
        Read = 0,
        Unread = 1,
        Hide = 2
    }
}
