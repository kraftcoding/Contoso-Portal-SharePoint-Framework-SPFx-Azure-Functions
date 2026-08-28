using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;

namespace Contoso.Portal.Data.DTO.Event;

public class EventAgendaItemDTO : BaseEventJsonFileDTO<EventAgendaItemDAO, EventAgendaItemJsonDAO>
{
    public string Description { get => ItemDAO.DocumentSetDescription; set => ItemDAO.DocumentSetDescription = value; }
    public double Duration { get => ItemDAO.Duracion; set => ItemDAO.Duracion = value; }
    public double Order { get => ItemDAO.Ordenacion; set => ItemDAO.Ordenacion = value; }

    public string? AgendaItemTypeId
    {
        get => ItemDAO.TipoAgendaItem?.TermId.ToString(); set
        {
            ItemDAO.TipoAgendaItem = value.AsTaxonomyFieldValue();
        }
    }

    public string[] RelatedDocumentsIds { get => JsonDAO.RelatedDocumentsIds; set => JsonDAO.RelatedDocumentsIds = value; }

    public string ParentId { get => ItemDAO.ParentId; set => ItemDAO.ParentId = value; }

    public string? OrderTypeId
    {
        get => ItemDAO.OrderType?.TermId.ToString(); set
        {
            ItemDAO.OrderType = value.AsTaxonomyFieldValue();
        }
    }

    // if agreemnt data is eCNTy, use agenda item data
    public string AgreementTitle { get => string.IsNullOrECNTy(JsonDAO.AgreementTitle) ? ItemDAO.Title : JsonDAO.AgreementTitle; set => JsonDAO.AgreementTitle = value; }
    public string AgreementDescription { get => string.IsNullOrECNTy(JsonDAO.AgreementDescription) ? ItemDAO.DocumentSetDescription : JsonDAO.AgreementDescription; set => JsonDAO.AgreementDescription = value; }
    public string AgreementRelatedCertificateId { get => JsonDAO.AgreementRelatedCertificateId; set => JsonDAO.AgreementRelatedCertificateId = value; }

    public string? AgreementStatusId
    {
        get => ItemDAO.EstadoAcuerdo?.TermId.ToString(); set
        {
            ItemDAO.EstadoAcuerdo = value.AsTaxonomyFieldValue();
        }
    }

    public string? OrderFormatted { get; internal set; } //  calculated
}