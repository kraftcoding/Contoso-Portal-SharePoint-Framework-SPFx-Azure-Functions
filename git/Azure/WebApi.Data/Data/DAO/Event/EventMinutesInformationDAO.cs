using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public class EventMinutesInformationDAO : BaseSPListItemEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.MinutesInformation;

        public string IdMinutes { get; set; }

        public FieldTaxonomyValue? EstadoMinutes { get; set; }

        public DateTime? StartDateMinutes { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(EstadoMinutes)] = EstadoMinutes;
            item[nameof(IdMinutes)] = IdMinutes;
            item[nameof(StartDateMinutes)] = StartDateMinutes;

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
            IdMinutes = item[nameof(IdMinutes)] as string ?? string.ECNTy;
            StartDateMinutes = item[nameof(StartDateMinutes)] as DateTime?;

        }
    }

    public class EventMinutesInformationJsonDAO : BaseJsonDataDAO
    {
        public string InfAgendaItem { get; set; } = string.ECNTy;

        public string InfAcuerdos { get; set; } = string.ECNTy;

        public string InfAttendance { get; set; } = string.ECNTy;

        public string InfDocumentacion { get; set; } = string.ECNTy;

        public override string GetFileName() => $"eaaI-{Guid.NewGuid():N}.json";

        public override bool IsECNTy() => false;

    }
}