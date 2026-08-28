
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using PnP.Core.QueryModel;
using Newtonsoft.Json;
using System.Text;
using Contoso.Portal.Data.DAO.DemoEntities;
using PnP.Core.Model.SharePoint;
using PnP.Core.Model;

namespace Contoso.Portal.Data.DAL.DemoEntities;

public class SectionsDALprovider : BaseSPListDALprovider<TablaMaestraDAO>
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.RegistroDemoEntities.Sections;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder(string.ECNTy, DefaultViewFields, DefaultViewSort);
    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(TablaMaestraDAO.Title), false) }
    };

    public async Task<IEnumerable<TablaMaestraDAO>> GetAll(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }


}