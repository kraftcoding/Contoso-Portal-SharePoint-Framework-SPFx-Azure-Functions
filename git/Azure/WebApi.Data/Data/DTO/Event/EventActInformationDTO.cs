using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;

namespace Contoso.Portal.Data.DTO.Event;

public class EventMinutesInformationDTO : BaseEventJsonFileDTO<EventMinutesInformationDAO, EventMinutesInformationJsonDAO>
{
    public string IdMinutes { get => ItemDAO.IdMinutes; set => ItemDAO.IdMinutes = value; }
    public DateTime? StartDateMinutes { get => ItemDAO.StartDateMinutes; set => ItemDAO.StartDateMinutes = value; }
    public string? EstadoMinutes { get => ItemDAO.EstadoMinutes?.TermId.ToString(); set => ItemDAO.EstadoMinutes = value.AsTaxonomyFieldValue(); }

    public string InfAgendaItem { get => JsonDAO.InfAgendaItem; set => JsonDAO.InfAgendaItem = value; }
    public string InfAcuerdos { get => JsonDAO.InfAcuerdos; set => JsonDAO.InfAcuerdos = value; }
    public string InfAttendance { get => JsonDAO.InfAttendance; set => JsonDAO.InfAttendance = value; }
    public string InfDocumentacion { get => JsonDAO.InfDocumentacion; set => JsonDAO.InfDocumentacion = value; }
}
