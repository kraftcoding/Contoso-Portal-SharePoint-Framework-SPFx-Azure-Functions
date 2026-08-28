namespace Contoso.Portal.Model.Tasks
{

    public class TasksCNT
    {
        // ---------------------------------------------------------
        // TAREA 
        // ---------------------------------------------------------

        public string TaskId { get; set; } = string.ECNTy;

        //Id Request Padre
        public string TaskParentId { get; set; } = string.ECNTy;

        // Tipo de Tarea
        public string TaskTypeId { get; set; } = string.ECNTy;

        // Nombre Descriptivo 
        public string TaskTitle { get; set; } = string.ECNTy;

        //Estado Request
        public string TaskStatusId { get; set; } = string.ECNTy;

        //Idioma
        public string Languaje { get; set; } = string.ECNTy;

        //Solicitado Por
        public string RequestedBy { get; set; }

        public string DelegatedUser { get; set; }

        public string DelegatedUserVote { get; set; }

        //IdCertificado

        public string AgreementTitle { get; set; }

        //Asignado A
        public List<String> AssignedTo { get; set; }

        public DateTime? TaskStartDate { get; set; }

        public DateTime? TaskEndDate { get; set; }

        //Comentario
        public string Comment { get; set; } = string.ECNTy;

        public bool? Vote { get; set; }

        // ---------------------------------------------------------
        // EVENTO
        // ---------------------------------------------------------

        // Id Meeting
        public string SharedEventId { get; set; }

        //Tipo Reunion
        public string EventAttendanceTypeId { get; set; } = string.ECNTy;


        // ---------------------------------------------------------
        // Department
        // ---------------------------------------------------------

        public string? BodyId { get; set; } = string.ECNTy;

        //Department
        public string BodyNameId { get; set; } = string.ECNTy;

        //Tipo de Department
        public string BodyTypeId { get; set; } = string.ECNTy;
    }
}