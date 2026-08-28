using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Management;
using Contoso.Portal.Data.DTO.Management;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants.Management;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Management;

public class RequestDALprovider() : BaseSPListDALprovider<RequestDAO>()
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.ManagementRequests;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } =
    [
        (nameof(BaseSPListItemDAO.Created), false)
    ];

    public async Task<IEnumerable<RequestDAO>> GetNewOrInprodgressRequestsByOperation(IPnPContext ctx, string operation)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetOperatorQuery(nameof(RequestDTO.OperacionPeticion), "Text", operation, ComparisonOperators.Eq)}<Or>{CAML_GetOperatorQuery(nameof(RequestDTO.EstadoPeticion), "Text", RequestStatus.New, ComparisonOperators.Eq)}{CAML_GetOperatorQuery(nameof(RequestDTO.EstadoPeticion), "Text", RequestStatus.Inprodgress, ComparisonOperators.Eq)}</Or></And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }

    public async Task<IEnumerable<RequestDAO>> GetRequestsByOperation(IPnPContext ctx, string operation)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetOperatorQuery(nameof(RequestDTO.OperacionPeticion), "Text", operation, ComparisonOperators.Eq)}</Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }

    public async Task<IEnumerable<RequestDAO>> GetDoneRequestsByOperation(IPnPContext ctx, string operation)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetOperatorQuery(nameof(RequestDTO.OperacionPeticion), "Text", operation, ComparisonOperators.Eq)}<Or>{CAML_GetOperatorQuery(nameof(RequestDTO.EstadoPeticion), "Text", RequestStatus.New, ComparisonOperators.Eq)}{CAML_GetOperatorQuery(nameof(RequestDTO.EstadoPeticion), "Text", RequestStatus.Inprodgress, ComparisonOperators.Eq)}</Or></And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }

    // §6: Consulta Request de asignacion por UPN + BodyId + Operation
    public async Task<IEnumerable<RequestDAO>> GetRequestByUpnBodyAndOperation(
        IPnPContext ctx, string upn, string bodyId, string operation)
    {
        var where =
            $"<Where><And>" +
            $"{CAML_GetOperatorQuery(nameof(RequestDTO.OperacionPeticion), "Text", operation, ComparisonOperators.Eq)}" +
            $"<And>" +
            $"{CAML_GetOperatorQuery(nameof(RequestDTO.UpnPeticion), "Text", upn, ComparisonOperators.Eq)}" +
            $"{CAML_GetOperatorQuery(nameof(RequestDTO.DepartmentPeticion), "Text", bodyId, ComparisonOperators.Eq)}" +
            $"</And></And></Where>";
        var view = PnPContentHelpers.CamlViewBuilder(where, DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }
}