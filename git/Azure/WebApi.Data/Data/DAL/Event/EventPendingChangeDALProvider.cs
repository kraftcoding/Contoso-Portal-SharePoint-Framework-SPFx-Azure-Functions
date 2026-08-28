using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Model.Events;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Event
{
    public class EventPendingChangeDALprovider
    {

        public List<PendingChanges> processEventChanges(EventChanges eventChanges)
        {
            var pendingChanges = new List<PendingChanges>();

            AddChange(pendingChanges, Fields.Title, eventChanges.Title);
            AddChange(pendingChanges, Fields.Description, eventChanges.Description);
            AddDateChange(pendingChanges, Fields.StartDate, eventChanges.StartDate);
            AddDateChange(pendingChanges, Fields.EndDate, eventChanges.EndDate);
            AddChange(pendingChanges, Fields.EventType, eventChanges.MeetingType?.TermId.ToString());
            AddChange(pendingChanges, Fields.EventTool, eventChanges.OnlineTool?.TermId.ToString());
            AddChange(pendingChanges, Fields.Location, eventChanges.Location);
            AddChange(pendingChanges, Fields.LocationDetails, eventChanges.LocationDetails);
            AddChange(pendingChanges, Fields.EventToolUrl, eventChanges.UrlOnlineTool?.Url);

            return pendingChanges;
        }

        private static void AddChange(List<PendingChanges> pendingChanges, string typeItem, string value)
        {
            if (value != null)
            {
                pendingChanges.Add(new PendingChanges
                {
                    ChangeType = typeItem,
                    ChangeValue = value
                });
            }
        }

        private static void AddDateChange(List<PendingChanges> pendingChanges, string typeItem, DateTime? date)
        {
            if (date != null)
            {
                pendingChanges.Add(new PendingChanges
                {
                    ChangeType = typeItem,
                    ChangeValue = date.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:sszzz")
                });
            }
        }

        public static List<PendingChanges> FilterAttendance(List<AttendanceChanges> attendanceChanges)
        {
            var pendingChanges = new List<PendingChanges>();
            pendingChanges.AddRange(attendanceChanges.Select(attendance => new PendingChanges { ChangeType = attendance.AddOrRemoved, ChangeValue = attendance.user }));
            return pendingChanges;
        }

        public static List<PendingChanges> FilterAddRemoveUpdateDocs(List<DocSetDocumentInformation> documentsChange)
        {
            var pendingChanges = new List<PendingChanges>();

            if (documentsChange.Count > 0)
            {
                foreach (var doc in documentsChange)
                {
                    pendingChanges.Add(new PendingChanges
                    {
                        ChangeType = GetTypeItemFromContentTypeId(doc.ContentTypeId),
                        ChangeValue = GetTypeItemTitle(doc.ContentTypeId, doc.ItemTitle, doc.FileName)
                    });
                }
            }

            return pendingChanges;
        }

        private static string GetTypeItemTitle(string contentTypeId, string itemTitle, string fileName)
        {
            return contentTypeId switch
            {
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.AgendaItem) || ctId.StartsWith(DALConstants.ContentTypeIds.Acuerdo) => itemTitle,
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.DocumentEnConstruccion) || ctId.StartsWith(DALConstants.ContentTypeIds.MinutesDocument) || ctId.StartsWith(DALConstants.ContentTypeIds.MinutesInformation) => fileName,
                _ => string.ECNTy
            };
        }

        private static string GetTypeItemFromContentTypeId(string contentTypeId)
        {
            return contentTypeId switch
            {
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.MinutesInformation) => "Minutes",
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.MinutesDocument) => "Minutes",
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.AgendaItem) => "AgendaItem",
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.DocumentEnConstruccion) => "Document",
                var ctId when ctId.StartsWith(DALConstants.ContentTypeIds.Acuerdo) => "Agreement",
                _ => string.ECNTy
            };
        }

        public EventChanges FilterEventChanges(BaseEventDAO currentVersion, string[] changedFields, BaseEventDAO publishedVersion)
        {
            var eventChanges = new EventChanges
            {
                Title = changedFields.Contains(nameof(MeetingEnEdicionDAO.Title)) ? currentVersion.Title : null,
                Description = (changedFields.Contains(nameof(MeetingEnEdicionDAO.DocumentSetDescription)) && !currentVersion.DocumentSetDescription.Equals(publishedVersion.DocumentSetDescription)) ? currentVersion.DocumentSetDescription : null,
                StartDate = changedFields.Contains(nameof(MeetingEnEdicionDAO.StartDate)) ? currentVersion.StartDate : null,
                EndDate = changedFields.Contains(nameof(MeetingEnEdicionDAO.EndDate)) ? currentVersion.EndDate : null,
                MeetingType = changedFields.Contains(nameof(MeetingEnEdicionDAO.MeetingType)) ? currentVersion.MeetingType : null,
                Location = (changedFields.Contains(nameof(MeetingEnEdicionDAO.Location)) && !currentVersion.Location.Equals(publishedVersion.Location)) ? currentVersion.Location : null,
                LocationDetails = (changedFields.Contains(nameof(MeetingEnEdicionDAO.LocationDetails)) && !currentVersion.LocationDetails.Equals(publishedVersion.LocationDetails)) ? currentVersion.LocationDetails : null,
                OnlineTool = changedFields.Contains(nameof(MeetingEnEdicionDAO.OnlineTool)) ? currentVersion.OnlineTool : null,
                UrlOnlineTool = changedFields.Contains(nameof(MeetingEnEdicionDAO.UrlOnlineTool)) ? currentVersion.UrlOnlineTool : null
            };
            return eventChanges;
        }
    }
}