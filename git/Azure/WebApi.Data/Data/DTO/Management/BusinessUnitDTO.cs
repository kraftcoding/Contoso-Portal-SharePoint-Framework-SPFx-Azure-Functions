using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Management;

namespace Contoso.Portal.Data.DTO.Management;

public class BusinessUnitDTO
{
    public string? BusinessArea { get; set; }
    public string? Division { get; set; }
    public DateTime? StartDate { get; set; }
    public DateTime? EndDate { get; set; }
}