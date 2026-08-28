using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.DemoEntities
{
    public class DemoEntityDAO : BaseSPListItemDAO
    {
        // He asumido un nombre de constante para el ContentType, cámbialo si es necesario
        public override string GetContentTypeIdValue() => ContentTypeIds.DemoEntity;

        // prodpiedades
        public string? Folder_EMD { get; set; }
        public FieldLookupValue? Section_EMD { get; set; }
        public FieldLookupValue? ScopeTerritorial_EMD { get; set; }
        public FieldLookupValue? CancellationReason_EMD { get; set; }
        public DateTime? FechaConstitucion_EMD { get; set; }
        public DateTime? FechaInscripcion_EMD { get; set; }
        public string? DatosInscripcion_EMD { get; set; }
        public string? DomicilioSocial_EMD { get; set; }
        public string? NIF_EMD { get; set; }
        public string? Objeto_EMD { get; set; }
        public string? EntidadesIntegrantes_EMD { get; set; }
        public string? IdentidadTitulares_EMD { get; set; }
        public string? RetestsentanteNombre_EMD { get; set; }
        public string? RetestsentantePrimerApellido_EMD { get; set; }
        public string? RetestsentanteSegundoApellido_EMD { get; set; }
        public DateTime? RetestsentanteFechaNacimiento_EMD { get; set; }
        public string? RetestsentanteNacionalidad_EMD { get; set; }
        public string? RetestsentanteTipoDocument_EMD { get; set; }
        public string? RetestsentanteID_EMD { get; set; }
        public string? RetestsentanteCargo_EMD { get; set; }
        public string? NumExpedienteACCEDA_EMD { get; set; }
        public string? SolicitanteNombre_EMD { get; set; }
        public string? SolicitanteApellido_EMD { get; set; }
        public string? SolicitanteSegundoApellido_EMD { get; set; }
        public string? SolicitanteTipoDocument_EMD { get; set; }
        public string? SolicitanteNumIdentificacion_EMD { get; set; }
        public string? ContactoEmail_EMD { get; set; }
        public string? ContactoTelefono_EMD { get; set; }
        public string? NumExpedienteCancelacion_EMD { get; set; }
        public string? Estado_EMD { get; set; }
        public string? Asiento_EMD { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(Folder_EMD)] = Folder_EMD;
            item[nameof(Section_EMD)] = Section_EMD;
            item[nameof(FechaConstitucion_EMD)] = FechaConstitucion_EMD;
            item[nameof(FechaInscripcion_EMD)] = FechaInscripcion_EMD;
            item[nameof(DatosInscripcion_EMD)] = DatosInscripcion_EMD;
            item[nameof(DomicilioSocial_EMD)] = DomicilioSocial_EMD;
            item[nameof(NIF_EMD)] = NIF_EMD;
            item[nameof(Objeto_EMD)] = Objeto_EMD;
            item[nameof(EntidadesIntegrantes_EMD)] = EntidadesIntegrantes_EMD;
            item[nameof(IdentidadTitulares_EMD)] = IdentidadTitulares_EMD;
            item[nameof(RetestsentanteNombre_EMD)] = RetestsentanteNombre_EMD;
            item[nameof(RetestsentantePrimerApellido_EMD)] = RetestsentantePrimerApellido_EMD;
            item[nameof(RetestsentanteSegundoApellido_EMD)] = RetestsentanteSegundoApellido_EMD;
            item[nameof(RetestsentanteFechaNacimiento_EMD)] = RetestsentanteFechaNacimiento_EMD;
            item[nameof(RetestsentanteNacionalidad_EMD)] = RetestsentanteNacionalidad_EMD;
            item[nameof(RetestsentanteTipoDocument_EMD)] = RetestsentanteTipoDocument_EMD;
            item[nameof(RetestsentanteID_EMD)] = RetestsentanteID_EMD;
            item[nameof(RetestsentanteCargo_EMD)] = RetestsentanteCargo_EMD;
            item[nameof(NumExpedienteACCEDA_EMD)] = NumExpedienteACCEDA_EMD;
            item[nameof(SolicitanteNombre_EMD)] = SolicitanteNombre_EMD;
            item[nameof(SolicitanteApellido_EMD)] = SolicitanteApellido_EMD;
            item[nameof(SolicitanteSegundoApellido_EMD)] = SolicitanteSegundoApellido_EMD;
            item[nameof(SolicitanteTipoDocument_EMD)] = SolicitanteTipoDocument_EMD;
            item[nameof(SolicitanteNumIdentificacion_EMD)] = SolicitanteNumIdentificacion_EMD;
            item[nameof(ContactoEmail_EMD)] = ContactoEmail_EMD;
            item[nameof(ContactoTelefono_EMD)] = ContactoTelefono_EMD;
            item[nameof(CancellationReason_EMD)] = CancellationReason_EMD;
            item[nameof(NumExpedienteCancelacion_EMD)] = NumExpedienteCancelacion_EMD;
            item[nameof(Estado_EMD)] = Estado_EMD;
            item[nameof(Asiento_EMD)] = Asiento_EMD;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            Folder_EMD = item[nameof(Folder_EMD)] as string ?? string.ECNTy;
            Section_EMD = item[nameof(Section_EMD)] as FieldLookupValue;
            ScopeTerritorial_EMD = item[nameof(ScopeTerritorial_EMD)] as FieldLookupValue;
            CancellationReason_EMD = item[nameof(CancellationReason_EMD)] as FieldLookupValue;
            FechaConstitucion_EMD = item[nameof(FechaConstitucion_EMD)] as DateTime?;
            FechaInscripcion_EMD = item[nameof(FechaInscripcion_EMD)] as DateTime?;
            DatosInscripcion_EMD = item[nameof(DatosInscripcion_EMD)] as string ?? string.ECNTy;
            DomicilioSocial_EMD = item[nameof(DomicilioSocial_EMD)] as string ?? string.ECNTy;
            NIF_EMD = item[nameof(NIF_EMD)] as string ?? string.ECNTy;
            Objeto_EMD = item[nameof(Objeto_EMD)] as string ?? string.ECNTy;
            EntidadesIntegrantes_EMD = item[nameof(EntidadesIntegrantes_EMD)] as string ?? string.ECNTy;
            IdentidadTitulares_EMD = item[nameof(IdentidadTitulares_EMD)] as string ?? string.ECNTy;
            RetestsentanteNombre_EMD = item[nameof(RetestsentanteNombre_EMD)] as string ?? string.ECNTy;
            RetestsentantePrimerApellido_EMD = item[nameof(RetestsentantePrimerApellido_EMD)] as string ?? string.ECNTy;
            RetestsentanteSegundoApellido_EMD = item[nameof(RetestsentanteSegundoApellido_EMD)] as string ?? string.ECNTy;
            RetestsentanteFechaNacimiento_EMD = item[nameof(RetestsentanteFechaNacimiento_EMD)] as DateTime?;
            RetestsentanteNacionalidad_EMD = item[nameof(RetestsentanteNacionalidad_EMD)] as string ?? string.ECNTy;
            RetestsentanteTipoDocument_EMD = item[nameof(RetestsentanteTipoDocument_EMD)] as string ?? string.ECNTy;
            RetestsentanteID_EMD = item[nameof(RetestsentanteID_EMD)] as string ?? string.ECNTy;
            RetestsentanteCargo_EMD = item[nameof(RetestsentanteCargo_EMD)] as string ?? string.ECNTy;
            NumExpedienteACCEDA_EMD = item[nameof(NumExpedienteACCEDA_EMD)] as string ?? string.ECNTy;
            SolicitanteNombre_EMD = item[nameof(SolicitanteNombre_EMD)] as string ?? string.ECNTy;
            SolicitanteApellido_EMD = item[nameof(SolicitanteApellido_EMD)] as string ?? string.ECNTy;
            SolicitanteSegundoApellido_EMD = item[nameof(SolicitanteSegundoApellido_EMD)] as string ?? string.ECNTy;
            SolicitanteTipoDocument_EMD = item[nameof(SolicitanteTipoDocument_EMD)] as string ?? string.ECNTy;
            SolicitanteNumIdentificacion_EMD = item[nameof(SolicitanteNumIdentificacion_EMD)] as string ?? string.ECNTy;
            ContactoEmail_EMD = item[nameof(ContactoEmail_EMD)] as string ?? string.ECNTy;
            ContactoTelefono_EMD = item[nameof(ContactoTelefono_EMD)] as string ?? string.ECNTy;
            NumExpedienteCancelacion_EMD = item[nameof(NumExpedienteCancelacion_EMD)] as string ?? string.ECNTy;
            Estado_EMD = item[nameof(Estado_EMD)] as string ?? string.ECNTy;
            Asiento_EMD = item[nameof(Asiento_EMD)] as string ?? string.ECNTy;
        }
    }
}