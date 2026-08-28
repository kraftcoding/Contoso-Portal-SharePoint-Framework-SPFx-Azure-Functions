
namespace Contoso.Portal.Model.Events
{
    public class EventMinutesInformation
    {
        public string AgendaInformation { get; set; } = string.ECNTy;
        public string AgreementsInformation { get; set; } = string.ECNTy;
        public string AttendanceInformation { get; set; } = string.ECNTy;
        public string DocumentationInformation { get; set; } = string.ECNTy;
        public string IdMinutes { get; set; } = string.ECNTy;
        public string MinutesStatus { get; set; } = string.ECNTy;
        public DateTime? StartDateMinutes { get; set; }
        public IEnumerable<EventUserAttendance> Attendances { get; set; } = [];

    }
}