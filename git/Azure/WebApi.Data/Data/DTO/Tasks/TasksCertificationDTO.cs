using Microsoft.Graph.Models;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DAO.Tasks;
using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DTO.Tasks
{
    public class TasksCertificationDTO : JsonFileListItemAbstractionDTO<TasksCertificationDAO, TasksCertificationJsonDAO>
    {

        // ---------------------------------------------------------
        // TAREA 
        // ---------------------------------------------------------

        public string Title { get => ItemDAO.Title; set => ItemDAO.Title = value; }

        public string? EstadoRequest { get => ItemDAO.EstadoRequest?.TermId.ToString(); set => ItemDAO.EstadoRequest = value.AsTaxonomyFieldValue(); }

        public string? Idioma { get => ItemDAO.Idioma?.TermId.ToString(); set => ItemDAO.Idioma = value.AsTaxonomyFieldValue(); }

        public FieldUserValue[] AsignadoA { get => ItemDAO.AsignadoA; set => ItemDAO.AsignadoA = value; }

        public FieldUserValue SolicitadoPor { get => ItemDAO.SolicitadoPor; set => ItemDAO.SolicitadoPor = value; }

        public DateTime? RequestStartDate { get => ItemDAO.RequestStartDate; set => ItemDAO.RequestStartDate = value; }

        public DateTime? RequestEndDate { get => ItemDAO.RequestEndDate; set => ItemDAO.RequestEndDate = value; }

        public string Comentario { get => JsonDAO.Comentario; set => JsonDAO.Comentario = value; }

        public string AcuerdoAsociado { get => ItemDAO.AcuerdoAsociado; set => ItemDAO.AcuerdoAsociado = value; }

        public string AgreementSharedId { get => ItemDAO.Acuerdo; set => ItemDAO.Acuerdo = value; }

        // ---------------------------------------------------------
        // EVENTO
        // ---------------------------------------------------------
        public string? IdMeeting { get => ItemDAO.IdMeeting; set => ItemDAO.IdMeeting = value; }

        // ---------------------------------------------------------
        // Department
        // ---------------------------------------------------------

        public string? Department { get => ItemDAO.Department?.TermId.ToString(); set => ItemDAO.Department = value.AsTaxonomyFieldValue(); }

        public string? TipoDepartment { get => ItemDAO.TipoDepartment.TermId.ToString(); set => ItemDAO.TipoDepartment = value.AsTaxonomyFieldValue(); }

    }
}