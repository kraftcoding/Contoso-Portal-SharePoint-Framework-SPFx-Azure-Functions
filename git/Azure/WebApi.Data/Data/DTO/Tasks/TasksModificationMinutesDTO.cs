using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Tasks;
using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DTO.Tasks
{
    public class TasksModificationMinutesDTO : JsonFileListItemAbstractionDTO<TasksModificationMinutesDAO, TasksModificationMinutesJsonDAO>
    {

        // ---------------------------------------------------------
        // TAREA 
        // ---------------------------------------------------------

        public string Title { get => ItemDAO.Title; set => ItemDAO.Title = value; }

        public string EstadoRequest { get => ItemDAO.EstadoRequest.TermId.ToString(); set => ItemDAO.EstadoRequest = value.AsTaxonomyFieldValue(); }

        public string IdRequestPadre { get => ItemDAO.IdRequestPadre; set => ItemDAO.IdRequestPadre = value; }

        public string IdMinutes { get => ItemDAO.IdMinutes; set => ItemDAO.IdMinutes = value; }

        public string Idioma { get => ItemDAO.Idioma.TermId.ToString(); set => ItemDAO.Idioma = value.AsTaxonomyFieldValue(); }

        public FieldUserValue? SolicitadoPor { get => ItemDAO.SolicitadoPor; set => ItemDAO.SolicitadoPor = value; }

        public FieldUserValue[]? AsignadoA { get => ItemDAO.AsignadoA; set => ItemDAO.AsignadoA = value; }

        public DateTime? RequestStartDate { get => ItemDAO.RequestStartDate; set => ItemDAO.RequestStartDate = value; }

        public DateTime? RequestEndDate { get => ItemDAO.RequestEndDate; set => ItemDAO.RequestEndDate = value; }

        public string Comentario { get => JsonDAO.Comentario; set => JsonDAO.Comentario = value; }

        // ---------------------------------------------------------
        // EVENTO
        // ---------------------------------------------------------

        public string? IdMeeting { get => ItemDAO.IdMeeting; set => ItemDAO.IdMeeting = value; }

        // ---------------------------------------------------------
        // Department
        // ---------------------------------------------------------

        public string Department { get => ItemDAO.Department.TermId.ToString(); set => ItemDAO.Department = value.AsTaxonomyFieldValue(); }

        public string TipoDepartment { get => ItemDAO.TipoDepartment.TermId.ToString(); set => ItemDAO.TipoDepartment = value.AsTaxonomyFieldValue(); }


    }
}