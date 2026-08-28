using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;

namespace Contoso.Portal.Model.Notifications
{
    public class NotificationMessage
    {
        public string Title { get; set; }

        public string Body { get; set; }

        public string TypeMessage { get; set; }

        public string PriorityMessage { get; set; }

        public bool SendMail{get; set;}
    }
    
}