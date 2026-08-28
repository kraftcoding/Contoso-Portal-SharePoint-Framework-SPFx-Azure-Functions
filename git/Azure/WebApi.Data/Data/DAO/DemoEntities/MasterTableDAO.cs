using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.DemoEntities
{
    public class TablaMaestraDAO : BaseSPListItemDAO
    {
        // He asumido un nombre de constante para el ContentType, cámbialo si es necesario
        public override string GetContentTypeIdValue() => ContentTypeIds.DemoEntity;

        // prodpiedades
        public string? Descripcion { get; set; }
        public bool Activo { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(Descripcion)] = Descripcion;
            item[nameof(Activo)] = Activo;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            Descripcion = item[nameof(Descripcion)] as string ?? string.ECNTy;
            Activo = item[nameof(Activo)] as bool? ?? false;
        }
    }
}