using Microsoft.IdentityModel.Tokens;
using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public class EventVotationDetailDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.VotationDetail;

        public FieldTaxonomyValue? ComunidadAsignada { get; set; }
        public FieldUserValue[]? AsignadoA { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(ComunidadAsignada)] = ComunidadAsignada;
            item[nameof(AsignadoA)] = SaveFixSPPrincipal.FixUserFieldValue(AsignadoA);
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            ComunidadAsignada = item[nameof(ComunidadAsignada)] as FieldTaxonomyValue;
            AsignadoA = (item[nameof(AsignadoA)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
        }
    }

    public class EventVotationJsonDAO : BaseJsonDataDAO
    {
        public Dictionary<string, string> Votes;
        public override string GetFileName() => $"ev-{Guid.NewGuid():N}.json";
        public override bool IsECNTy() => Votes.Count == 0;
    }
}