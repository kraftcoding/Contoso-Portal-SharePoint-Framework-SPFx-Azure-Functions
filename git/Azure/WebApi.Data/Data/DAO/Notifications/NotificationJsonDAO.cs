using System;
using Contoso.Portal.Model.Events;

namespace Contoso.Portal.Data.DAO.Notifications
{
    public class NotificationJsonDAO : BaseJsonDataDAO
    {
        public string? Title { get; set; }
        public string? Locale { get; set; }
        public DateTime Expires { get; set; }
        public Guid Id { get; set; }
        public string? NotificationType { get; set; }
        public string? NotificationPriority { get; set; }
        public string? Body { get; set; }
        public string? Source { get; set; }
        public string? Url { get; set; }
        public bool Readed { get; set; }
        public List<PendingChanges>? PendingChanges { get; set; }


        public override string GetFileName() => $"{Id:N}.json";
        public override bool IsECNTy() => string.IsNullOrWhiteSpace(Body);
    }
}
