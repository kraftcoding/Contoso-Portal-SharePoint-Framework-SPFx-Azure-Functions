using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Tasks
{
    public class TasksDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.Task.Request;

        public FieldTaxonomyValue? EstadoRequest { get; set; }

        public FieldTaxonomyValue? FormatoAttendance { get; set; }
        public DateTime? EndDate { get; set; }
        public DateTime? StartDate { get; set; }
        public FieldUserValue[]? AsignadoA { get; set; }

        public FieldUserValue? SolicitadoPor { get; set; }
        public FieldUserValue? Delegado { get; set; }

        public FieldUserValue? DelegadoVoto { get; set; }
        public Boolean? VotoDelegado { get; set; }

        public string IdRequestPadre { get; set; } = string.ECNTy;

        public bool? Attendance { get; set; }
        public string AcuerdoAsociado { get; set; } = string.ECNTy;
        public bool? Votar { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(IdMeeting)] = IdMeeting;
            item[nameof(EstadoRequest)] = EstadoRequest;
            item[nameof(Department)] = Department;
            item[nameof(FormatoAttendance)] = FormatoAttendance;
            item[nameof(EndDate)] = EndDate;
            item[nameof(StartDate)] = StartDate;
            item[nameof(Attendance)] = Attendance;
            item[nameof(IdRequestPadre)] = IdRequestPadre;
            item[nameof(AcuerdoAsociado)] = AcuerdoAsociado;
            item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
            item[nameof(Delegado)] = SaveFixSPPrincipal.FixUserFieldValue(Delegado);
            item[nameof(DelegadoVoto)] = SaveFixSPPrincipal.FixUserFieldValue(DelegadoVoto);
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);
            item[nameof(VotoDelegado)] = VotoDelegado;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            ID = (int)item[nameof(ID)];
            Title = item[nameof(Title)] as string ?? string.ECNTy;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            ContentTypeId = item[nameof(ContentTypeId)] as string ?? string.ECNTy;
            IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;
            AcuerdoAsociado = item[nameof(AcuerdoAsociado)] as string ?? string.ECNTy;
            EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;
            FormatoAttendance = item[nameof(FormatoAttendance)] as FieldTaxonomyValue;
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            EndDate = item[nameof(EndDate)] as DateTime?;
            StartDate = item[nameof(StartDate)] as DateTime?;
            Attendance = (bool)item[nameof(Attendance)];
            Votar = (bool)item[nameof(Votar)];
            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;

            //TODO: Revisar Task

            //Solo Tareas Delegadas
            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestDelegation))
                Delegado = item[nameof(Delegado)] as FieldUserValue;

            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestDelegation))
                DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;

            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestDelegation))
                VotoDelegado = (bool)item[nameof(VotoDelegado)];


            //Solo Tareas Attendance
            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestAttendance))
                Delegado = item[nameof(Delegado)] as FieldUserValue;

            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestAttendance))
                DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;

            if ((item[nameof(ContentTypeId)] as string).Contains(DALConstants.ContentTypeIds.Task.RequestAttendance))
                VotoDelegado = (bool)item[nameof(VotoDelegado)];

        }
    }
    public class TasksJsonDAO : BaseJsonDataDAO
    {
        public override string GetFileName() => $"eadr-{Guid.NewGuid():N}.json";

        public string Comentario { get; set; } = string.ECNTy;
        public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
    }
}