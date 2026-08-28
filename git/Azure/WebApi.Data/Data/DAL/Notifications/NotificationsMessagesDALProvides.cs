using System.Text;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Data.Extensions;
using Newtonsoft.Json;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using PnP.Core.QueryModel;
using System.Diagnostics;
using PnP.Core.Model.SharePoint;


namespace Contoso.Portal.Data.Data.DAL.Notifications
{
    public class NotificationsMessagesDALprovides : BaseSPListDALprovider<NotificationMessagesDAO>
    {
        public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.NotificationConfig;

        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

        public async Task<NotificationMessagesJsonDAO> GetMessagesConfig(IPnPContext ctx)
        {
            try
            {
                var result = new NotificationMessagesJsonDAO();

                var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

                var viewXml = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery("Activo", "Boolean", "1", ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);

                await list.LoadItemsByCamlQueryAsync(new CamlQueryOptions()
                {
                    ViewXml = viewXml,
                    DatesInUtc = true
                }, p => p.RoleAssignments.Queryproperties(p => p.PrincipalId, p => p.RoleDefinitions));

                var idMessage = 0;
                foreach (var listItem in list.Items.AsRequested())
                    idMessage = listItem.Id;

                if (idMessage != 0)
                    return await GetMessagesDAO(ctx, idMessage);

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting GetMessages", ex);
            }
        }

        private async Task<NotificationMessagesJsonDAO> GetMessagesDAO(IPnPContext ctx, int idNotification)
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));

            var messageListItem = await list.Items.GetByIdAsync(idNotification, li => li.All, li => li.File);

            var fileBytes = await messageListItem.File.GetContentBytesAsync();

            return JsonConvert.DeserializeObject<NotificationMessagesJsonDAO>(Encoding.UTF8.GetString(fileBytes));

        }

        private string GetDebuggerDisplay()
        {
            return ToString();
        }
    }

}
