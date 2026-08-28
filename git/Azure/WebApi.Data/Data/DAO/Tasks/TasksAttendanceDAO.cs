using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksAttendanceDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestAttendance;

        public FieldUserValue? Delegado { get; set; }

        public FieldUserValue? DelegadoVoto { get; set; }

        public bool? VotoDelegado { get; set; }

        public string IdRequestPadre { get; set; } = string.ECNTy;

        public FieldUserValue[]? AsignadoA { get; set; }
        public FieldUserValue? SolicitadoPor { get; set; }
        public FieldTaxonomyValue? FormatoAttendance { get; set; }
        public FieldTaxonomyValue? EstadoRequest { get; set; }

        public DateTime? RequestStartDate { get; set; }

        public DateTime? RequestEndDate { get; set; }

        public Boolean Votar { get; set; }

        public Boolean Attendance { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoRequest)] = EstadoRequest;
            item[nameof(FormatoAttendance)] = FormatoAttendance;
            item[nameof(Votar)] = Votar;
            item[nameof(IdRequestPadre)] = IdRequestPadre;

            // item[nameof(RequestStartDate)] = RequestStartDate;
            // item[nameof(RequestEndDate)] = RequestEndDate;
            item[nameof(Attendance)] = Attendance;
            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);
            item[nameof(Delegado)] = SaveFixSPPrincipal.FixUserFieldValue(Delegado);
            item[nameof(DelegadoVoto)] = SaveFixSPPrincipal.FixUserFieldValue(DelegadoVoto);
            item[nameof(VotoDelegado)] = VotoDelegado;

            // Campos que se autorrellenan del DocumentSet
            //Idioma 
            //IdMeeting
            //Tipo de Reunión
            //Department
            //Tipo de Department
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;
            FormatoAttendance = item[nameof(FormatoAttendance)] as FieldTaxonomyValue;
            Votar = (bool)item[nameof(Votar)];
            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            // RequestStartDate = item[nameof(RequestStartDate)] is not null ? (DateTime)item[nameof(RequestStartDate)] : throw new Exception($"{nameof(RequestStartDate)} is null in item {item.Id}");           
            // RequestEndDate = item[nameof(RequestEndDate)] is not null ? (DateTime)item[nameof(RequestEndDate)] : throw new Exception($"{nameof(RequestEndDate)} is null in item {item.Id}");  
            Attendance = (bool)item[nameof(Attendance)];
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;
            MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;
            TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;

            Delegado = item[nameof(Delegado)] as FieldUserValue;
            DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;
            VotoDelegado = (bool)item[nameof(VotoDelegado)];
        }
    }

    public class TasksAttendanceJsonDAO : BaseJsonDataDAO
    {
        public string Comentario { get; set; } = string.ECNTy;

        public override string GetFileName() => $"earAtt-{Guid.NewGuid():N}.json";
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}