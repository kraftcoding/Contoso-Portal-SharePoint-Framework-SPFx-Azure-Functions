using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Management;

public class RequestDAO : BaseSPListItemDAO
{
    public override string GetContentTypeIdValue() => ContentTypeIds.ManagementRequest;

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

    public override void MapToSPListItem(IListItem item)
    {
        base.MapToSPListItem(item);

        item[nameof(UpnPeticion)]        = UpnPeticion;
        item[nameof(DepartmentPeticion)]     = DepartmentPeticion;
        item[nameof(OperacionPeticion)]  = OperacionPeticion;
        item[nameof(ParametrosPeticion)] = ParametrosPeticion;
        item[nameof(EstadoPeticion)]     = EstadoPeticion;
        item[nameof(NumIntentos)]        = NumIntentos;
        item[nameof(ErrorMensaje)]       = ErrorMensaje;
        item[nameof(ErrorPermanente)]    = ErrorPermanente;
        item[nameof(IdOperation)]        = IdOperation;
    }

    public override void MapFromSPListItem(IListItem item)
    {
        base.MapFromSPListItem(item);

        UpnPeticion        = item[nameof(UpnPeticion)]        as string ?? string.ECNTy;
        DepartmentPeticion     = item[nameof(DepartmentPeticion)]     as string ?? string.ECNTy;
        OperacionPeticion  = item[nameof(OperacionPeticion)]  as string ?? string.ECNTy;
        ParametrosPeticion = item[nameof(ParametrosPeticion)] as string ?? string.ECNTy;
        EstadoPeticion     = item[nameof(EstadoPeticion)]     as string ?? string.ECNTy;
        NumIntentos        = item[nameof(NumIntentos)] is int i ? i : 0;
        ErrorMensaje       = item[nameof(ErrorMensaje)]    as string ?? string.ECNTy;
        ErrorPermanente    = item[nameof(ErrorPermanente)] is bool b && b;
        IdOperation        = item[nameof(IdOperation)]     as string ?? string.ECNTy;
    }

    public override Dictionary<string, object> AsNewListItem()
    {
        var values = base.AsNewListItem();
        values[nameof(UpnPeticion)]        = UpnPeticion;
        values[nameof(DepartmentPeticion)]     = DepartmentPeticion;
        values[nameof(OperacionPeticion)]  = OperacionPeticion;
        values[nameof(ParametrosPeticion)] = ParametrosPeticion;
        values[nameof(EstadoPeticion)]     = EstadoPeticion;
        values[nameof(NumIntentos)]        = NumIntentos;
        values[nameof(ErrorMensaje)]       = ErrorMensaje;
        values[nameof(ErrorPermanente)]    = ErrorPermanente;
        values[nameof(IdOperation)]        = IdOperation;
        return values;
    }
}
