using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Model.Events
{
    public class PendingChanges
    {
        public string ChangeType { get; set; } = string.ECNTy;
        public string ChangeValue { get; set; } = string.ECNTy;
    }

    public class AttendanceChanges
    {
        public string AddOrRemoved { get; set; } = string.ECNTy;
        public string user { get; set; } = string.ECNTy;
    }

    public class AttendeesChangesResult
    {
        public List<string> Added { get; set; }
        public List<string> Removed { get; set; }
        public List<string> NotChanged { get; set; }
    }

    public class EventChanges
    {
        public string Title { get; set; } = string.ECNTy;
        public string Description { get; set; } = string.ECNTy;
        public DateTime? StartDate { get; set; }

        public DateTime? EndDate { get; set; }

        public FieldTaxonomyValue? MeetingType { get; set; }
        public string Location { get; set; } = string.ECNTy;
        public string LocationDetails { get; set; } = string.ECNTy;
        public FieldTaxonomyValue? OnlineTool { get; set; }
        public FieldUrlValue? UrlOnlineTool { get; set; }
    }

    public class PendingChangesBool
    {
        public bool meetingPointsTabHasChanges { get; set; }
        public bool meetingAssistantsTabHasChanges { get; set; }
        public bool meetingDocumentsTabHasChanges { get; set; }
        public bool meetingAgreeementsTabHasChanges { get; set; }
        public bool meetingMinutesTabHasChanges { get; set; }
        public bool meetingDetailsHasChanges { get; set; }
    }
}
