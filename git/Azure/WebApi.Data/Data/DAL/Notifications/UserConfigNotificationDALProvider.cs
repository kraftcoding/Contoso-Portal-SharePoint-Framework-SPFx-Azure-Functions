using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Data.DTO.Notifications;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.Data.DAL.Notifications
{
    public class UserConfigNotificationDALprovider : BaseSPListDALproviderForJsonFiles<NotificationUserConfigDAO, NotificationUserConfigJsonDAO, NotificationUserConfigDTO>
    {
        public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.NotificationConfig;
        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
        public async Task<IEnumerable<NotificationUserConfigDTO?>> GetUserNotificationConfig(IPnPContext ctx, string userId)
        {
            NotificationUserConfigDTO dto = new NotificationUserConfigDTO();
            dto.JsonDAO.userId = userId;
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                var fileNameQuery = CAML_GetOperatorQuery("FileLeafRef", "File", dto.JsonDAO.GetFileName(), ComparisonOperators.Eq);
                var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{fileNameQuery}{CAML_GetContentTypeFilterPart()}</And></Where>", DefaultViewFields, DefaultViewSort);
                return await GetDTOByView(notificationConext, view);
            }
        }

        public async Task UpdateUserNotificationConfig(IPnPContext ctx, string userId, NotificationUserConfigDTO newValue)
        {
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                var result = await AddOrUpdateInLibrary(notificationConext, new List<NotificationUserConfigDTO>() { newValue });
            }
        }

        public async Task<NotificationUserConfigDTO?> AddDefaultConfig(IPnPContext ctx, string userId)
        {
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                var dto = new NotificationUserConfigDTO();
                dto.JsonDAO.userId = userId;
                var result = await AddOrUpdateInLibrary(notificationConext, new List<NotificationUserConfigDTO>() { dto });
                return result.FirstOrDefault();
            }
        }
    }
}
