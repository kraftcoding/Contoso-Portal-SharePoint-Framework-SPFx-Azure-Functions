
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using PnP.Core.QueryModel;
using Newtonsoft.Json;
using System.Text;

namespace Contoso.Portal.Data.DAL.ConfigDepartments;

public class ConfigDepartmentsDALprovider : BaseSPListDALprovider<ConfigDepartmentsDAO>
{
    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.Home.ConfigDepartments;
    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(ConfigDepartmentsDAO.FechaConstitucion), false) }
    };

    public async Task<IEnumerable<ConfigDepartmentsDAO>> GetBodies(IPnPContext ctx)
    {
        return await GetByView(ctx, DefaultView);
    }

    public async Task<ConfigDepartmentsDAO> GetBodyById(IPnPContext ctx, string bodyId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(ConfigDepartmentsDAO.Title), "Text", bodyId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
        return (await GetByView(ctx, view)).FirstOrDefault();
    }

    public async Task<IEnumerable<ConfigDepartmentsDAO>> GetBodiesById(IPnPContext ctx, string[] bodyIds)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetInOperatorQuery(nameof(ConfigDepartmentsDAO.Title), "Text", bodyIds.ToList())}</And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }

    public async Task<ConfigEmailBodiesJsonDAO> GetEmailBodies(IPnPContext ctx, int idEmailBodies)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync(idEmailBodies, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            var result = JsonConvert.DeserializeObject<ConfigEmailBodiesJsonDAO>(Encoding.UTF8.GetString(fileBytes));

            return result;

        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting EmailBodies from id: {idEmailBodies}", ex);
        }
    }

    public async Task<Stream> GetCertificateTemplateBodies(IPnPContext ctx, int idCertificateBody)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync(idCertificateBody, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            return new MemoryStream(fileBytes);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetCertificateTemplateBodies from id: {idCertificateBody}", ex);
        }
    }


    public async Task<Stream> GetAgreementTemplateBodies(IPnPContext ctx, int idAgreement)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync(idAgreement, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            return new MemoryStream(fileBytes);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error  GetAgreemtTemplateBodies from id: {idAgreement}", ex);
        }
    }

    public async Task<Stream> GetAgendaTemplateBodies(IPnPContext ctx, int idAgenda)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync(idAgenda, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            MemoryStream stream = new MemoryStream(fileBytes);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error  GetAgreemtTemplateBodies from id: {idAgenda}", ex);
        }
    }

    public async Task<Stream> GetAttendanceTemplateBodies(IPnPContext ctx, int idAttendanceBodies)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var emailListItem = await list.Items.GetByIdAsync(idAttendanceBodies, li => li.All, li => li.File);

            var fileBytes = await emailListItem.File.GetContentBytesAsync();

            return new MemoryStream(fileBytes);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetAgreemtTemplateBodies from id: {idAttendanceBodies}", ex);
        }
    }
}