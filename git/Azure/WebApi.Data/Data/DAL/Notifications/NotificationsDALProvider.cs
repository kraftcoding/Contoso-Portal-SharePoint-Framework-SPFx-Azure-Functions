using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Data.DTO.Notifications;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL.Notifications
{
    public class NotificationsDALprovider : BaseSPListDALproviderForJsonFiles<NotificationsDAO, NotificationJsonDAO, NotificationDTO>
    {
        public NotificationsDALprovider(bool isHidden)
        {
            this.isHiddenNotification = isHidden;
        }

        public bool isHiddenNotification { get; private set; }
        public override string ListSiteRelavteUrl => this.isHiddenNotification ? ListsSiteRelativeUrls.NotificationsReaded : ListsSiteRelativeUrls.NotificationsToRead;
        public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);
        public async Task<IEnumerable<NotificationDTO>> GetUserNotifications(IPnPContext ctx, string userFolder, string locale, string? source)
        {
            var localeQuery = CAML_GetOperatorQuery(DALConstants.Fields.NotificationEntity.Locale, "Text", locale, ComparisonOperators.Eq);
            var sourceQuery = CAML_GetOperatorQuery(DALConstants.Fields.NotificationEntity.Source, "Text", source, ComparisonOperators.Eq);

            var metadataQuery = string.IsNullOrWhiteSpace(source) ? localeQuery : $"<And>{sourceQuery}{localeQuery}</And>";

            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{metadataQuery}</And></Where>", DefaultViewFields, DefaultViewSort);
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                return await GetDTOByView(notificationConext, view, notificationConext.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(userFolder));
            }
        }

        public async Task<IEnumerable<NotificationDTO>> GetArchivedNotifications(IPnPContext ctx, DateTime dateFilter)
        {
            var dateFilterPart = CAML_GetOperatorQuery(nameof(BaseEventDAO.Created), dateFilter, ComparisonOperators.Leq);
            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{dateFilterPart}</And></Where>", DefaultViewFields, DefaultViewSort);
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                return await GetDTOByView(notificationConext, view, notificationConext.Uri.AbsolutePath.UriCombine(ListsSiteRelativeUrls.NotificationsReaded));
            }
        }

        public async Task<IEnumerable<NotificationDTO>> GetUserArchivedNotifications(IPnPContext ctx, string userFolder, string locale, string? source)
        {
            var localeQuery = CAML_GetOperatorQuery(DALConstants.Fields.NotificationEntity.Locale, "Text", locale, ComparisonOperators.Eq);
            var sourceQuery = CAML_GetOperatorQuery(DALConstants.Fields.NotificationEntity.Source, "Text", source, ComparisonOperators.Eq);

            var metadataQuery = string.IsNullOrWhiteSpace(source) ? localeQuery : $"<And>{sourceQuery}{localeQuery}</And>";
            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{metadataQuery}</And></Where>", DefaultViewFields, DefaultViewSort);
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                return await GetDTOByView(notificationConext, view, notificationConext.Uri.AbsolutePath.UriCombine(ListsSiteRelativeUrls.NotificationsReaded).UriCombine(userFolder));
            }
        }

        public async Task<NotificationDTO?> GetNotificationById(IPnPContext ctx, string path, Guid id)
        {
            var fileNameQuery = CAML_GetOperatorQuery("FileLeafRef", "File", $"{id:N}.json", ComparisonOperators.Eq);
            var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{fileNameQuery}{CAML_GetContentTypeFilterPart()}</And></Where>", DefaultViewFields, DefaultViewSort);
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
                return (await base.GetDTOByView(notificationConext, view, notificationConext.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(path))).FirstOrDefault();
        }

        public async Task RemoveNotifcation(IPnPContext ctx, string path, NotificationDTO notification)
        {
            await notification.ItemDAO.AsListItem().DeleteAsync();
        }

        public async Task AddNotifcation(IPnPContext notificationConext, string path, NotificationDTO notification)
        {
            IFolder folder = await notificationConext.Web.GetFolderByServerRelativeUrlAsync(notificationConext.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(path), p => p.ServerRelativeUrl);
            await AddOrUpdateInFolder(notificationConext, folder, [notification]);
        }

        public async Task UpdateNotification(IPnPContext notificationConext, string path, NotificationDTO notification)
        {
            IFolder folder = await notificationConext.Web.GetFolderByServerRelativeUrlAsync(notificationConext.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl).UriCombine(path), p => p.ServerRelativeUrl);
            await AddOrUpdateInFolder(notificationConext, folder, [notification]);
        }

        public async Task<Guid> CreateUserConfigStructure(IPnPContext ctx)
        {
            using (var notificationConext = await ctx.CloneAsync(DALConstants.PnPContexts.RootSiteAsSystem))
            {
                var guid = Guid.NewGuid();
                var list = await notificationConext.Web.Lists.GetByServerRelativeUrlAsync(notificationConext.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl));
                await list.AddListFolderAsync($"{guid:N}");
                return guid;
            }
        }
    }
}