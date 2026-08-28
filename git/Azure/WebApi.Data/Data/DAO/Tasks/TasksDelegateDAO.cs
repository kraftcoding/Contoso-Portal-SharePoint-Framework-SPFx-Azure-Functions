using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksDelegateDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestDelegation;

        public string IdRequestPadre { get; set; } = string.ECNTy;

        public FieldTaxonomyValue EstadoRequest { get; set; }

        public FieldUserValue? SolicitadoPor { get; set; }

        public FieldUserValue[]? AsignadoA { get; set; }

        public FieldUserValue? Delegado { get; set; }

        public FieldUserValue? DelegadoVoto { get; set; }

        public bool VotoDelegado { get; set; }

        public DateTime? RequestStartDate { get; set; }

        public DateTime? RequestEndDate { get; set; }

        public FieldTaxonomyValue MeetingType { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoRequest)] = EstadoRequest;
            item[nameof(IdRequestPadre)] = IdRequestPadre;
            item[nameof(VotoDelegado)] = VotoDelegado;

            // if(RequestStartDate != null)    
            //     item[nameof(RequestStartDate)] = RequestStartDate;

            // if(RequestEndDate != null)    
            //     item[nameof(RequestEndDate)] = RequestEndDate;


            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(Delegado)] = SaveFixSPPrincipal.FixUserFieldValue(Delegado);
            item[nameof(DelegadoVoto)] = SaveFixSPPrincipal.FixUserFieldValue(DelegadoVoto);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);

            // Campos que se autorrellenan del DocumentSet
            //Idioma 
            //IdMeeting
            //Tipo de Reunión
            //Estado Meeting
            //Department
            //Tipo de Department             
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;

            VotoDelegado = (bool)item[nameof(VotoDelegado)];
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;
            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            Delegado = item[nameof(Delegado)] as FieldUserValue;
            DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;
            Idioma = item[nameof(Idioma)] as FieldTaxonomyValue;
            MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;

            //TODO: Corregir
            //RequestStartDate = item[nameof(RequestStartDate)] is not null ? (DateTime)item[nameof(RequestStartDate)] : throw new Exception($"{nameof(RequestStartDate)} is null in item {item.Id}");           
            //RequestEndDate = item[nameof(RequestEndDate)] is not null ? (DateTime)item[nameof(RequestEndDate)] : throw new Exception($"{nameof(RequestEndDate)} is null in item {item.Id}");        
        }
    }

    public class TasksDelegateJsonDAO : BaseJsonDataDAO
    {
        public override string GetFileName() => $"earDel-{Guid.NewGuid():N}.json";

        public string Comentario { get; set; } = string.ECNTy;
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}