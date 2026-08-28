using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Management
{
    public class processReportsDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.processReport;

        public string? Tipoprodceso { get; set; }
        public string? EstadoPeticion { get; set; }
        public DateTime? Inicioprodceso { get; set; }
        public DateTime? Finprodceso { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(Tipoprodceso)] = Tipoprodceso;
            item[nameof(EstadoPeticion)] = EstadoPeticion;
            item[nameof(Inicioprodceso)] = Inicioprodceso;
            item[nameof(Finprodceso)] = Finprodceso;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);

            Tipoprodceso = item[nameof(Tipoprodceso)] as string ?? string.ECNTy;
            EstadoPeticion = item[nameof(EstadoPeticion)] as string ?? string.ECNTy;
            Inicioprodceso = item[nameof(Inicioprodceso)] is not null ? (DateTime)item[nameof(Inicioprodceso)] : null;
            Finprodceso = item[nameof(Finprodceso)] is not null ? (DateTime)item[nameof(Finprodceso)] : null;

        }

        public override Dictionary<string, object> AsNewListItem()
        {
            var values = base.AsNewListItem();
            values[nameof(Tipoprodceso)] = Tipoprodceso;
            values[nameof(EstadoPeticion)] = EstadoPeticion;
            values[nameof(Inicioprodceso)] = Inicioprodceso;
            values[nameof(Finprodceso)] = Finprodceso;
            return values;
        }
    }
}