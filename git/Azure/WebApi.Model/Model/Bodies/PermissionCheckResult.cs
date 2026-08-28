namespace Contoso.Portal.Model.Bodies;

// §6: Resultado enriquecido de CheckPermissions.
// Incluye pertenencia al grupo Y estado de la Request en RequestesAdministracion.
public class PermissionCheckResult
{
    public bool    HasPermissions    { get; set; }
    public string  AssignmentStatus  { get; set; } = "None";
    public int     NumIntentos       { get; set; }
    public bool    IsErrorPermanente { get; set; }
    public string? ErrorMensaje      { get; set; }
    public string? IdOperation       { get; set; }
}
