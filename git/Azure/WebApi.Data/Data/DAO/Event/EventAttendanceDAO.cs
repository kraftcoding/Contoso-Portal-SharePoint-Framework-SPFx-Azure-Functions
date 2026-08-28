using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAL.Helpers;

namespace Contoso.Portal.Data.DAO.Event;

public class EventAttendanceDAO : BaseSPListItemEventDAO
{
    public override string GetContentTypeIdValue() => ContentTypeIds.Task.RequestAttendance;

    public FieldUserValue[]? AsignadoA { get; set; }
    public FieldUserValue? SolicitadoPor { get; set; }
    public FieldTaxonomyValue? FormatoAttendance { get; set; }
    public FieldTaxonomyValue? EstadoRequest { get; set; }
    // public DateTime? RequestStartDate { get; set; }
    // public DateTime? RequestEndDate { get; set; }
    public bool Attendance { get; set; }

    public string IdRequestPadre { get; set; } = string.ECNTy;

    public bool Votar { get; set; }
    public bool? VotoDelegado { get; set; }

    public FieldUserValue? Delegado { get; set; }
    public FieldUserValue? DelegadoVoto { get; set; }

    public override void MapToSPListItem(IListItem item)
    {
        base.MapToSPListItem(item);
        base.MapToSPListItemEvent(item);

        item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);
        item[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
        item[nameof(FormatoAttendance)] = FormatoAttendance;
        item[nameof(EstadoRequest)] = EstadoRequest;

        item[nameof(Attendance)] = Attendance;
        item[nameof(IdRequestPadre)] = IdRequestPadre;

        item[nameof(Votar)] = Votar;
        item[nameof(VotoDelegado)] = VotoDelegado;

        item[nameof(Delegado)] = SaveFixSPPrincipal.FixUserFieldValue(Delegado);
        item[nameof(DelegadoVoto)] = SaveFixSPPrincipal.FixUserFieldValue(DelegadoVoto);
    }

    public override void MapFromSPListItem(IListItem item)
    {
        base.MapFromSPListItem(item);
        base.MapFromSPListItemEvent(item);

        AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
        SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;
        FormatoAttendance = item[nameof(FormatoAttendance)] as FieldTaxonomyValue;
        EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;

        Attendance = (bool)item[nameof(Attendance)];
        IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;

        Votar = (bool)item[nameof(Votar)];
        VotoDelegado = (bool)item[nameof(VotoDelegado)];

        Delegado = item[nameof(Delegado)] as FieldUserValue;
        DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;
    }

    public override void MapFromSPListItemVersion(IListItemVersion item)
    {
        base.MapFromSPListItemVersion(item);
        base.MapFromSPListItemEventVersion(item);

        AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
        SolicitadoPor = item[nameof(SolicitadoPor)] as FieldUserValue;
        FormatoAttendance = item[nameof(FormatoAttendance)] as FieldTaxonomyValue;
        EstadoRequest = item[nameof(EstadoRequest)] as FieldTaxonomyValue;

        Attendance = (bool)item[nameof(Attendance)];
        IdRequestPadre = item[nameof(IdRequestPadre)] as string ?? string.ECNTy;

        Votar = (bool)item[nameof(Votar)];
        VotoDelegado = (bool)item[nameof(VotoDelegado)];

        Delegado = item[nameof(Delegado)] as FieldUserValue;
        DelegadoVoto = item[nameof(DelegadoVoto)] as FieldUserValue;
    }

    public override Dictionary<string, object> AsNewListItem()
    {
        var values = base.AsNewListItemEvent(base.AsNewListItem());

        values[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);
        values[nameof(SolicitadoPor)] = SaveFixSPPrincipal.FixUserFieldValue(SolicitadoPor);
        values[nameof(FormatoAttendance)] = FormatoAttendance;
        values[nameof(EstadoRequest)] = EstadoRequest;

        values[nameof(Attendance)] = Attendance;
        values[nameof(IdRequestPadre)] = IdRequestPadre;

        values[nameof(Votar)] = Votar;
        values[nameof(VotoDelegado)] = VotoDelegado;

        values[nameof(Delegado)] = SaveFixSPPrincipal.FixUserFieldValue(Delegado);
        values[nameof(DelegadoVoto)] = SaveFixSPPrincipal.FixUserFieldValue(DelegadoVoto);

        return (Dictionary<string, object>)values;
    }
}

public class EventAttendanceJsonDAO : BaseJsonDataDAO
{
    public string Comentario { get; set; } = string.ECNTy;

    public override string GetFileName() => $"ear-{Guid.NewGuid():N}.json";
    public override bool IsECNTy() => string.IsNullOrECNTy(Comentario);
}