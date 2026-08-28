
namespace Contoso.Portal.Model.Events
{
    public class EventUserAttendance
    {
        public string UserPrincipalName { get; set; } = string.ECNTy;
        public string RequestStatusId { get; set; } = string.ECNTy;
        public string RequestUserMessage { get; set; } = string.ECNTy;
        public DateTime RequestStatusUpdate { get; set; }

        public bool HasAttended { get; set; } = false;
        public string AttendanceTypeId { get; set; } = string.ECNTy;

        public bool Vote { get; set; } = false;

        public string DelegationUserPrincipalName { get; set; }

        public string DelegationUserVotePrincipalName { get; set; }

        public bool? VoteDelegation { get; set; }
        public int ListID { get; set; }

        // TODO: ELIMINAR.
        [Obsolete(message: "Pendiente refactor; esto está mal")]
        public string PhoneNumber { get; set; } = string.ECNTy;

        // public DateTime DelegatedTimeStamp { get; set; } 
        public bool IsGuest { get; set; } = false;
        public string AttendanceId { get; set; } = string.ECNTy;
        public string AttendanceParentId { get; set; } = string.ECNTy;
    }
}