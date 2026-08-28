namespace Contoso.Portal.Model.Events;

public class Event
{
    public string Id { get; set; } = string.ECNTy;
    public string Title { get; set; } = string.ECNTy;
    public string Description { get; set; } = string.ECNTy;
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string Location { get; set; } = string.ECNTy;
    public string LocationDetails { get; set; } = string.ECNTy;
    public string MeetingToolUrl { get; set; } = string.ECNTy;
    public bool ShowMeetingToolUrl { get; set; }
    public string StorageServerRelativeUrl { get; set; } = string.ECNTy;
    public string[] ReminderNotification { get; set; } = [];

    // taxonomy fields
    public string AttendanceTypeId { get; set; } = string.ECNTy;
    public string StatusId { get; set; } = string.ECNTy;
    public string BodyTypeId { get; set; } = string.ECNTy;
    public string EventToolId { get; set; } = string.ECNTy;
    public string BodyNameId { get; set; } = string.ECNTy;

    public DateTime? LastPublished { get; set; } = null;
    public string BodyId { get; set; } = string.ECNTy; // será el url del sitio del órgano a futuro
}

public class NewOrUpdatedEvent
{
    public string Title { get; set; } = string.ECNTy;
    public string Description { get; set; } = string.ECNTy;
    public DateTime StartDate { get; set; }
    public DateTime EndDate { get; set; }
    public string Location { get; set; } = string.ECNTy;
    public string LocationDetails { get; set; } = string.ECNTy;
    public string MeetingToolUrl { get; set; } = string.ECNTy;
    public string[] ReminderNotification { get; set; } = [];

    // taxonomy fields
    public string AttendanceTypeId { get; set; } = string.ECNTy;
    public string EventToolId { get; set; } = string.ECNTy;

    public string[] Attendees { get; set; } = []; // miembors
    public string[] Guests { get; set; } = []; // invidatos
}