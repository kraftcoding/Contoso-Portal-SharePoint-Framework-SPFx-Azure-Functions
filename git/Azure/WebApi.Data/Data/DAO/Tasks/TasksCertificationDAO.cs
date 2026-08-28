using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksCertificationDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestCertificacion;

        public String AcuerdoAsociado { get; set; }

        public String Acuerdo { get; set; }

        public FieldUserValue[]? AsignadoA { get; set; }
        public FieldUserValue? SolicitadoPor { get; set; }
        public FieldTaxonomyValue? EstadoRequest { get; set; }

        public DateTime? RequestStartDate { get; set; }

        public DateTime? RequestEndDate { get; set; }

        public FieldTaxonomyValue? Idioma { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoRequest)] = EstadoRequest;
            item[nameof(IdMeeting)] = IdMeeting;
            item[nameof(AcuerdoAsociado)] = AcuerdoAsociado;
            item[nameof(Acuerdo)] = Acuerdo;

            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);

            item[nameof(RequestStartDate)] = RequestStartDate;
            item[nameof(RequestEndDate)] = RequestEndDate;

            // Campos que se autorrellenan del DocumentSet
            //Idioma 
            //IdMeeting
            //Department
            //Tipo de Department
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            AcuerdoAsociado = item[nameof(AcuerdoAsociado)] as string ?? string.ECNTy;
            Acuerdo = item[nameof(Acuerdo)] as string ?? string.ECNTy;

            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;

            RequestStartDate = item[nameof(RequestStartDate)] is not null ? (DateTime)item[nameof(RequestStartDate)] : null;
            RequestEndDate = item[nameof(RequestEndDate)] is not null ? (DateTime)item[nameof(RequestEndDate)] : null;
        }
    }

    public class TasksCertificationJsonDAO : BaseJsonDataDAO
    {
        public string Comentario { get; set; } = string.ECNTy;

        public override string GetFileName() => $"earCer-{Guid.NewGuid():N}.json";
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}