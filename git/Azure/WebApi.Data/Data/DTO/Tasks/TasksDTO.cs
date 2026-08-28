using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Tasks;
using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DTO.Tasks
{
    public class TasksDTO : JsonFileListItemAbstractionDTO<TasksDAO, TasksJsonDAO>
    {

        // ---------------------------------------------------------
        // TAREA 
        // ---------------------------------------------------------

        public string Title { get => ItemDAO.Title; set => ItemDAO.Title = value; }

        public string IdRequestPadre { get => ItemDAO.IdRequestPadre; set => ItemDAO.IdRequestPadre = value; }

        public string ContentTypeId { get => ItemDAO.ContentTypeId; set => ItemDAO.ContentTypeId = value; }

        public string? EstadoRequest { get => ItemDAO.EstadoRequest?.TermId.ToString(); set => ItemDAO.EstadoRequest = value.AsTaxonomyFieldValue(); }

        public Boolean Votar { get => (bool)ItemDAO.Votar; set => ItemDAO.Votar = value; }

        public Boolean Attendance { get => (bool)ItemDAO.Attendance; set => ItemDAO.Attendance = value; }

        public string? FormatoAttendance { get => ItemDAO.FormatoAttendance?.TermId.ToString(); set => ItemDAO.FormatoAttendance = value.AsTaxonomyFieldValue(); }

        public FieldUserValue[] AsignadoA { get => ItemDAO.AsignadoA; set => ItemDAO.AsignadoA = value; }

        public FieldUserValue SolicitadoPor { get => ItemDAO.SolicitadoPor; set => ItemDAO.SolicitadoPor = value; }

        public DateTime? StartDate { get => ItemDAO.StartDate; set => ItemDAO.StartDate = value; }

        public DateTime? EndDate { get => ItemDAO.EndDate; set => ItemDAO.EndDate = value; }

        public DateTime? Created { get => ItemDAO.Created; set => ItemDAO.Created = (DateTime)value; }

        public FieldUserValue Delegado { get => ItemDAO.Delegado; set => ItemDAO.Delegado = value; }

        public FieldUserValue DelegadoVoto { get => ItemDAO.DelegadoVoto; set => ItemDAO.DelegadoVoto = value; }

        public string AcuerdoAsociado { get => ItemDAO.AcuerdoAsociado; set => ItemDAO.AcuerdoAsociado = value; }

        public string Comentario { get => JsonDAO.Comentario; set => JsonDAO.Comentario = value; }

        public bool? VotoDelegado { get => ItemDAO.VotoDelegado; set => ItemDAO.VotoDelegado = value; }

        // ---------------------------------------------------------
        // EVENTO
        // ---------------------------------------------------------
        public string? IdMeeting { get => ItemDAO.IdMeeting; set => ItemDAO.IdMeeting = value; }

        // ---------------------------------------------------------
        // Department
        // ---------------------------------------------------------

        public string? Department { get => ItemDAO.Department?.TermId.ToString(); set => ItemDAO.Department = value.AsTaxonomyFieldValue(); }
    }
}