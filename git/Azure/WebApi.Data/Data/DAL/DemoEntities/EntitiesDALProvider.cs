
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

public class EntidadesDALprovider : BaseSPListDALprovider<DemoEntityDAO>
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.RegistroDemoEntities.Entidades;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
    public const string TemplateFolder = "PlantillasBase";
    public const string TemplateName = "PlantillaCertificadoEntidad";
    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(DemoEntityDAO.Title), false) }
    };

    public async Task<IEnumerable<DemoEntityDAO>> GetEntidades(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }

    public async Task<Stream> GetTemplate(IPnPContext ctx)
    {
        try
        {
            var folderRelativeURL = ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(TemplateFolder);
            var folder = await ctx.Web.GetFolderByServerRelativeUrlAsync(folderRelativeURL) ?? throw new Exception("'{TemplateFolder}' folder could not found");
            await folder.LoadAsync(x => x.Files);
            var file = folder.Files.FirstOrDefault(x => x.Name.StartsWith(TemplateName)) ?? throw new Exception($"Template '{TemplateName}' could not found");

            var fileBytes = await file.GetContentBytesAsync();
            return new MemoryStream(fileBytes);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetTemplate: {ex.Message}", ex);
        }
    }

    public async Task<DemoEntityDAO?> GetEntidadById(IPnPContext ctx, int id)
    {
        var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Id, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
        var item = await list.Items.GetByIdAsync(id,
            i => i.All,
            i => i.File
        );
        await item.LoadAsync(p => p.FieldValuesAsText);
        var field = list.Fields.FirstOrDefault(f => f.Title == "Section_EMD");
        var result = new DemoEntityDAO();
        result.MapFromSPListItem(item);
        return result;
    }

    /*public async Task<DemoEntityDAO?> Create(IPnPContext ctx, DemoEntityDAO bodyDao)
    {
        var (item, _, _) = await PnPContentHelpers.CreateDocumentSet(ctx, ListSiteRelavteUrl, bodyDao.CodigoDepartment!, DAOBaseContentTypeId, bodyDao.AsNewListItem());
        return await this.GetBodyById(ctx, item.Id);
    }*/
}