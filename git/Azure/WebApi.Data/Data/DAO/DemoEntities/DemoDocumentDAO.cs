using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.DemoEntities
{
    public class DocumentDAO : BaseSPListItemDAO
    {
        // He asumido un nombre de constante para el ContentType, cámbialo si es necesario
        public override string GetContentTypeIdValue() => ContentTypeIds.DocumentDemoEntity;

        // metadata
        public string? TipoDocument_EMD { get; set; }
        public string? Folder_EMD { get; set; }

        public string UniqueId { get; set; }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(TipoDocument_EMD)] = TipoDocument_EMD;
            item[nameof(Folder_EMD)] = Folder_EMD;
            item[nameof(UniqueId)] = UniqueId;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            TipoDocument_EMD = item[nameof(TipoDocument_EMD)] as string ?? string.ECNTy;
            Folder_EMD = item[nameof(Folder_EMD)] as string ?? string.ECNTy;
            UniqueId = item[nameof(UniqueId)].ToString();
        }
    }
}