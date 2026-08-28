using Contoso.Portal.Data.DAO;

namespace Contoso.Portal.Data.DTO;

public abstract class BaseEventJsonFileDTO<T, U>
        : JsonFileListItemAbstractionDTO<T, U>
        where T : BaseSPListItemEventDAO, new()
        where U : BaseJsonDataDAO, new()
{
    public string? LanguageId { get => ItemDAO.Idioma?.TermId.ToString(); }
    public string? BodyId { get => ItemDAO.Department?.TermId.ToString(); }
    public string? BodyTypeId { get => ItemDAO.TipoDepartment?.TermId.ToString(); }

    public string? SharedEventId { get => ItemDAO.IdMeeting; }
    public string? EventAttendanceTypeId { get => ItemDAO.MeetingType?.TermId.ToString(); }

    public string? UniqueSharedID { get => ItemDAO.UniqueSharedID; }
}