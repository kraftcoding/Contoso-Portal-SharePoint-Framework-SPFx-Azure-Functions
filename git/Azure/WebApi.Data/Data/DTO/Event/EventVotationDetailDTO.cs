using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DTO.Event;

public class EventVotationDetailDTO : BaseEventJsonFileDTO<EventVotationDetailDAO, EventVotationJsonDAO>
{
    public string? ComunidadAsignada { get => ItemDAO.ComunidadAsignada?.TermId.ToString(); set => ItemDAO.ComunidadAsignada = value.AsTaxonomyFieldValue(); }
    public Dictionary<string, string> Votes { get => JsonDAO.Votes; set => JsonDAO.Votes = value?.Count > 0 ? value : new Dictionary<string, string>(); }
    public FieldUserValue[]? AsignadoA { get => ItemDAO.AsignadoA; set => ItemDAO.AsignadoA = value; }
}
