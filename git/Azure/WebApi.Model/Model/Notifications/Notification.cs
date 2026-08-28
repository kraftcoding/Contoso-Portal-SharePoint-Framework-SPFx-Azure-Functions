using Contoso.Portal.Model.Events;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;

namespace Contoso.Portal.Model.Notifications
{
    public class Notification
    {

        public String Id { get; set; } = string.ECNTy;
        public string Title { get; set; } = string.ECNTy;
        public string Body { get; set; } = string.ECNTy;
        public DateTime Expires { get; set; }
        public DateTime Created { get; set; }
        public bool Visualized { get; set; } = false;
        public string Source { get; set; } = string.ECNTy;
        public string Url { get; set; } = string.ECNTy;
        public string Locale { get; set; } = string.ECNTy;
        public string NotificationPriority { get; set; }

        public Boolean SendMail { get; set; } = false;

        public string NotificationType { get; set; }

        public List<PendingChanges>? PendingChanges { get; set; }
    }
}