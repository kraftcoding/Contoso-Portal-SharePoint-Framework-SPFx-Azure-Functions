using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksModificationMinutesDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestModificacion;

        public string IdRequestPadre { get; set; } = string.ECNTy;

        public string IdMinutes { get; set; } = string.ECNTy;

        public FieldTaxonomyValue EstadoRequest { get; set; }

        public FieldUserValue SolicitadoPor { get; set; }

        public FieldUserValue[]? AsignadoA { get; set; }

        public DateTime? RequestStartDate { get; set; }

        public DateTime? RequestEndDate { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoRequest)] = EstadoRequest;
            item[nameof(IdRequestPadre)] = IdRequestPadre;
            item[nameof(IdMinutes)] = IdMinutes;

            item[nameof(RequestStartDate)] = RequestStartDate;
            item[nameof(RequestEndDate)] = RequestEndDate;

            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);

            // Campos que se autorrellenan del DocumentSet
            //Idioma 
            //IdMeeting
            //Estado Meeting
            //Department
            //Tipo de Department             
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;
            IdMinutes = item[nameof(IdMinutes)] as string ?? string.ECNTy;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;

            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;

            RequestStartDate = item[nameof(RequestStartDate)] is not null ? (DateTime)item[nameof(RequestStartDate)] : null;
            RequestEndDate = item[nameof(RequestEndDate)] is not null ? (DateTime)item[nameof(RequestEndDate)] : null;
        }
    }

    public class TasksModificationMinutesJsonDAO : BaseJsonDataDAO
    {
        public override string GetFileName() => $"earMod-{Guid.NewGuid():N}.json";

        public string Comentario { get; set; } = string.ECNTy;
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}