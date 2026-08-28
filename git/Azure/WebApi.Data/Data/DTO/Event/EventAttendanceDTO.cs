using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DTO.Event;

public class EventAttendanceDTO : BaseEventJsonFileDTO<EventAttendanceDAO, EventAttendanceJsonDAO>
{
    public FieldUserValue? RequestedBy
    {
        get => ItemDAO.SolicitadoPor;
        set => ItemDAO.SolicitadoPor = value;
    }

    public string? AttendanceTypeId // ¿Attendance del usuario?
    {
        get => ItemDAO.FormatoAttendance?.TermId.ToString();
        set => ItemDAO.FormatoAttendance = value.AsTaxonomyFieldValue();
    }

    public string? RequestStatus
    {
        get => ItemDAO.EstadoRequest?.TermId.ToString();
        set => ItemDAO.EstadoRequest = value.AsTaxonomyFieldValue();
    }
    public FieldUserValue[]? AssignedTo
    {
        get => ItemDAO.AsignadoA;
        set => ItemDAO.AsignadoA = value;
    }

    public bool CanVote
    {
        get => ItemDAO.Votar;
        set => ItemDAO.Votar = value;
    }

    public bool HasAttended // ¿Attendance del usuario?
    {
        get => ItemDAO.Attendance;
        set => ItemDAO.Attendance = value;
    }

    public string Comment { get => JsonDAO.Comentario; set => JsonDAO.Comentario = value; }
}