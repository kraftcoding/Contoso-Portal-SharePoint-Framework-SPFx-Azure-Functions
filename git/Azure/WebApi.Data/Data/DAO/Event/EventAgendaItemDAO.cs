using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event;

public class EventAgendaItemDAO : BaseSPListItemEventDAO
{
    public override string GetContentTypeIdValue() => ContentTypeIds.AgendaItem;

    public string DocumentSetDescription { get; set; } = string.ECNTy;
    public FieldTaxonomyValue? TipoAgendaItem { get; set; }
    public double Duracion { get; set; }
    public double Ordenacion { get; set; }
    public FieldTaxonomyValue? EstadoAcuerdo { get; set; }
    public FieldTaxonomyValue? OrderType { get; set; }
    public string ParentId { get; set; }

    public override void MapToSPListItem(IListItem item)
    {
        base.MapToSPListItem(item);
        base.MapToSPListItemEvent(item);

        item[nameof(DocumentSetDescription)] = DocumentSetDescription;
        item[nameof(TipoAgendaItem)] = TipoAgendaItem;
        item[nameof(Duracion)] = Duracion;
        item[nameof(Ordenacion)] = Ordenacion;
        item[nameof(EstadoAcuerdo)] = EstadoAcuerdo;
        item[nameof(OrderType)] = OrderType;
        item[nameof(ParentId)] = ParentId;
    }

    public override void MapFromSPListItem(IListItem item)
    {
        base.MapFromSPListItem(item);
        base.MapFromSPListItemEvent(item);

        DocumentSetDescription = item[nameof(DocumentSetDescription)] as string ?? string.ECNTy;
        TipoAgendaItem = item[nameof(TipoAgendaItem)] as FieldTaxonomyValue;
        Duracion = double.TryParse(item[nameof(Duracion)]?.ToString(), out var duracion) ? duracion : 0;
        Ordenacion = double.TryParse(item[nameof(Ordenacion)]?.ToString(), out var ordenacion) ? ordenacion : 0;
        EstadoAcuerdo = item[nameof(EstadoAcuerdo)] as FieldTaxonomyValue;
        OrderType = item[nameof(OrderType)] as FieldTaxonomyValue;
        ParentId = item[nameof(ParentId)] as string ?? string.ECNTy;
    }

    public override Dictionary<string, object> AsNewListItem()
    {
        var values = base.AsNewListItemEvent(base.AsNewListItem());

        values[nameof(DocumentSetDescription)] = DocumentSetDescription;
        values[nameof(TipoAgendaItem)] = TipoAgendaItem;
        values[nameof(Duracion)] = Duracion;
        values[nameof(Ordenacion)] = Ordenacion;
        values[nameof(EstadoAcuerdo)] = EstadoAcuerdo;
        values[nameof(OrderType)] = OrderType;
        values[nameof(ParentId)] = ParentId;

        return (Dictionary<string, object>)values;
    }
}

public class EventAgendaItemJsonDAO : BaseJsonDataDAO
{
    public string[] RelatedDocumentsIds { get; set; } = [];
    public string AgreementTitle { get; set; } = string.ECNTy;
    public string AgreementDescription { get; set; } = string.ECNTy;
    public string AgreementRelatedCertificateId { get; set; } = string.ECNTy;
    public string OrderFormatted { get; set; } = string.ECNTy;

    public override string GetFileName() => $"eai-{Guid.NewGuid():N}.json";
    public override bool IsECNTy() =>
        (RelatedDocumentsIds == null || RelatedDocumentsIds.Length == 0)
        && string.IsNullOrECNTy(AgreementTitle)
        && string.IsNullOrECNTy(AgreementDescription)
        && string.IsNullOrECNTy(AgreementRelatedCertificateId);
}