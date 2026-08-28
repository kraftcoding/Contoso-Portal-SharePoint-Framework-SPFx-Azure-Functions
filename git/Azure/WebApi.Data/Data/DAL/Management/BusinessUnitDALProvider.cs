using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Management;
using Contoso.Portal.Data.DTO.Management;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants.Management;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Management;

public class BusinessUnitDALprovider() : BaseSPListDALprovider<BusinessUnitDAO>()
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.BusinessUnit;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } =
    [
        (nameof(BaseSPListItemDAO.Created), false)
    ];

    public async Task<IEnumerable<BusinessUnitDAO>> GetBusinessAreasMinistries(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }

    public async Task<IEnumerable<BusinessUnitDAO>> GetBusinessAreasMinistriesByStartDate(IPnPContext ctx, DateTime startDate)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetOperatorQuery(nameof(BusinessUnitDAO.StartDate), "DateTime", startDate.ToString("yyyy-MM-ddTHH:mm:ssZ"), ComparisonOperators.Geq)}{CAML_GetOperatorQuery(nameof(BusinessUnitDAO.StartDate), "DateTime", startDate.AddDays(1).ToString("yyyy-MM-ddTHH:mm:ssZ"), ComparisonOperators.Lt)}</And>{CAML_GetContentTypeFilterPart()}</And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }
}