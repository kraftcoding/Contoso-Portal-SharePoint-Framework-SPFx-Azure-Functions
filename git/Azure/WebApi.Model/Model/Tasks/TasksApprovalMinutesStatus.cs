
using Contoso.Portal.Model.profile;

namespace Contoso.Portal.Model.Tasks
{
    public class TasksApprodvalMinutesStatus
    {

        public IEnumerable<TasksApprodvalMinutesStatusUserInfo> ApprodvalsPendings { get; set; }

        public IEnumerable<TasksApprodvalMinutesStatusUserInfo> ApprodvalsAccepted { get; set; }

        public IEnumerable<TasksApprodvalMinutesStatusUserInfo> ApprodvalsRejected { get; set; }

        public IEnumerable<TasksApprodvalMinutesStatusUserInfo> ApprodvalsPendingModification { get; set; }

        public string ApprodvalEndDate { get; set; }
    }

    public class TasksApprodvalMinutesStatusUserInfo
    {
        public string? Upn { get; set; }
        public string? FullName { get; set; }
    }
}