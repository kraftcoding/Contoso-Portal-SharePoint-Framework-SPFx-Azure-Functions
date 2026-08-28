
namespace Contoso.Portal.Model.Tasks
{
    public class TasksAttendance : TasksCNT
    {
        //Formato Attendance
        public string TasksAttendanceTypeId { get; set; }

        public bool? Vote { get; set; } = false;

        public bool? HasAttended { get; set; } = false;

        //Delegado a
        public string DelegateTo { get; set; }

        public string DelegatedUserVote { get; set; }

        public bool? DelegateVote { get; set; } = false;
    }
}