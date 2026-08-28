using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Management;

namespace Contoso.Portal.Data.DTO.Management;

public class RequestDTO
{
    public string UpnPeticion        { get; set; }
    public string DepartmentPeticion     { get; set; }
    public string OperacionPeticion  { get; set; }
    public string ParametrosPeticion { get; set; }
    public string EstadoPeticion     { get; set; }

    // §7: Nuevos campos requeridos en la lista RequestesAdministracion
    public int    NumIntentos     { get; set; }
    public string ErrorMensaje    { get; set; }
    public bool   ErrorPermanente { get; set; }
    public string IdOperation     { get; set; }
}