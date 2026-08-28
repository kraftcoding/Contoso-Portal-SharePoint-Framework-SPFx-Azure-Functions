
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using PnP.Core.QueryModel;
using Newtonsoft.Json;
using System.Text;

namespace Contoso.Portal.Data.DAL.ConfigDepartments;

public class ConfigEXTERNALBodiesDALprovider : BaseSPListDALprovider<ConfigEXTERNALBodiesDAO>
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.Home.ConfigEXTERNALBodies;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
    public const string TemplateFolder = "Plantillas Base";
    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(ConfigEXTERNALBodiesDAO.FechaConstitucion), false) }
    };

    public async Task<IEnumerable<ConfigEXTERNALBodiesDAO>> GetBodies(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }

    public async Task<ConfigEXTERNALBodiesDAO?> GetBodyById(IPnPContext ctx, int id)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(ConfigDepartmentsDAO.ID), "Text", id.ToString(), ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var result = (await GetByView(ctx, view)).FirstOrDefault();
        return result;
    }

    public async Task<int> GetTemplateID(IPnPContext ctx, string templateName)
    {
        try
        {
            var folderRelativeURL = ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(TemplateFolder);
            var folder = await ctx.Web.GetFolderByServerRelativeUrlAsync(folderRelativeURL) ?? throw new Exception("EXTERNAL bodies configuration folder could not found");
            await folder.LoadAsync(x => x.Files);
            var file = folder.Files.FirstOrDefault(x => x.Name.StartsWith(templateName)) ?? throw new Exception($"Template '{templateName}' could not found");
            await file.LoadAsync(x => x.ListItemAllFields);
            return (int)file.ListItemAllFields["ID"];
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetIDCertificateTemplateBody: {ex.Message}", ex);
        }
    }

    public async Task<Stream> GetCertificateTemplateBodies(IPnPContext ctx, int? idCertificateBody)
    {
        try
        {
            if (idCertificateBody is null)
                throw new ArgumentNullException($"Null CertificateBodyID column on body {idCertificateBody}");

            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync((int)idCertificateBody, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            return new MemoryStream(fileBytes);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetCertificateTemplateBodies from EXTERNAL bodyId: {idCertificateBody}", ex);
        }
    }

    public async Task<ConfigEXTERNALBodiesDAO?> Create(IPnPContext ctx, ConfigEXTERNALBodiesDAO bodyDao)
    {
        var (item, _, _) = await PnPContentHelpers.CreateDocumentSet(ctx, ListSiteRelavteUrl, bodyDao.CodigoDepartment!, DAOBaseContentTypeId, bodyDao.AsNewListItem());
        return await this.GetBodyById(ctx, item.Id);
    }
}