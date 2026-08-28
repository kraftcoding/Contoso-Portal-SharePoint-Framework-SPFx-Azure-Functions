
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

public class DocumentsDALprovider : BaseSPListDALprovider<DocumentDAO>
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.RegistroDemoEntities.Entidades;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
    public const string TemplateFolder = "Plantillas Base";
    public const string TemplateName = "PlantilaDemoEntity";
    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(DemoEntityDAO.Title), false) }
    };

    public async Task<IEnumerable<DocumentDAO>> GetAll(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }
}