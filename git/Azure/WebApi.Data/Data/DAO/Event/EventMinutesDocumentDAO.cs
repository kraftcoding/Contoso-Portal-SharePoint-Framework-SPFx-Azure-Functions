using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public class EventMinutesDocumentDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.MinutesDocument;

        public FieldTaxonomyValue? EstadoMinutes { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoMinutes)] = EstadoMinutes;

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

            EstadoMinutes = item[nameof(EstadoMinutes)] as FieldTaxonomyValue;
        }
    }
}