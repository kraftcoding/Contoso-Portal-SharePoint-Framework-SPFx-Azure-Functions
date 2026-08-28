using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Event;

namespace Contoso.Portal.Data.DTO.Event
{
    public class EventPublishedDocumentDTO<EventPublishedDocumentDAO>
    {
        public string ID { get; set; }
        public string Title { get; set; }
        public string FileRef { get; set; }
        public string FileLeafRef { get; set; }
        public string EventId { get; set; }
    }
}