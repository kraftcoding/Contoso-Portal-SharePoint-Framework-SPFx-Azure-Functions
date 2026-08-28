using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksApprodvalMinutesDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestAprodbacionMinutes;

        public string IdMinutes { get; set; } = string.ECNTy;

        public DateTime? StartDateMinutes { get; set; }


        public FieldUserValue[]? AsignadoA { get; set; }
        public FieldUserValue? SolicitadoPor { get; set; }
        public FieldTaxonomyValue? EstadoRequest { get; set; }

        public DateTime? RequestStartDate { get; set; }

        public DateTime? RequestEndDate { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);

            item[nameof(EstadoRequest)] = EstadoRequest;

            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);

            item[nameof(RequestStartDate)] = RequestStartDate;
            item[nameof(RequestEndDate)] = RequestEndDate;

            item[nameof(IdMinutes)] = IdMinutes;
            item[nameof(StartDateMinutes)] = StartDateMinutes;


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

            IdMinutes = item[nameof(IdMinutes)] as string ?? string.ECNTy;
            StartDateMinutes = item[nameof(StartDateMinutes)] is not null ? (DateTime)item[nameof(StartDateMinutes)] : null;
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;

            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;

            RequestStartDate = item[nameof(RequestStartDate)] is not null ? (DateTime)item[nameof(RequestStartDate)] : null;
            RequestEndDate = item[nameof(RequestEndDate)] is not null ? (DateTime)item[nameof(RequestEndDate)] : null;
        }
    }

    public class TasksApprodvalMinutesJsonDAO : BaseJsonDataDAO
    {
        public string Comentario { get; set; } = string.ECNTy;
        public override string GetFileName() => $"earApp-{Guid.NewGuid():N}.json";
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}