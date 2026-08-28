using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Management;
using Contoso.Portal.Data.DTO.Management;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants.Management;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Management;

public class processReportsDALprovider() : BaseSPListDALprovider<processReportsDAO>()
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.processReports;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } =
    [
        (nameof(BaseSPListItemDAO.Modified), false)
    ];

    public async Task<processReportsDAO> GetLastSuccesfulExecution(IPnPContext ctx)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(processReportsDAO.EstadoPeticion), "Text", "OK", ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
        //Ignora el pagesize, así que devolvemos filtrando por linq:
        var result = await GetByView(ctx, view);
        return result.ToList().OrderByDescending(x => x.Inicioprodceso).FirstOrDefault() ?? new processReportsDAO();
    }
}