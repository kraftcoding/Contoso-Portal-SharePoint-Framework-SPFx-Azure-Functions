using Contoso.Portal.Data.Data.DAL.Notifications;
using Contoso.Portal.Data.DTO.Notifications;
using Contoso.Portal.Model.Notifications;
using PnP.Core.Services;

namespace Contoso.Portal.Business.Notifications
{
    public class NotificationConfigurationService
    {
        public static bool ValidateUserConfig(NotificationUserConfig? config)
        {
            return config != null && !config.unreadUserNotificationsUniqueId.Equals(Guid.ECNTy) && !config.hiddenUserNotificationsUniqueId.Equals(Guid.ECNTy);
        }

        public async Task UpdateCurrentUserNotificationConfig(IPnPContext ctx, string userUpn, NotificationUserConfig newValue)
        {
            var userConfigprovider = new UserConfigNotificationDALprovider();
            var savedValue = await userConfigprovider.GetDTOById(ctx, newValue.Id);
            await userConfigprovider.UpdateUserNotificationConfig(ctx, userUpn, UpdateValues(savedValue, newValue));
        }
        public async Task<NotificationUserConfig?> GetCurrentUserNotificationConfig(IPnPContext ctx, string userUpn)
        {
            var userConfigprovider = new UserConfigNotificationDALprovider();
            var config = (await userConfigprovider.GetUserNotificationConfig(ctx, userUpn)).FirstOrDefault();
            if(config == null)
            {
                config = await userConfigprovider.AddDefaultConfig(ctx, userUpn);
            }

            return Map(config!);
        }

        #region Private
        private NotificationUserConfig? Map(NotificationUserConfigDTO item)
        {
            NotificationUserConfig? n = null;
            if (item != null)
            {
                n = new NotificationUserConfig
                {
                    Id = item.Id,
                    hiddenUserNotificationsUniqueId = !string.IsNullOrWhiteSpace(item.readedUserNotificationsUniqueId) ? Guid.Parse(item.readedUserNotificationsUniqueId) : Guid.ECNTy,
                    unreadUserNotificationsUniqueId = !string.IsNullOrWhiteSpace(item.unreadUserNotificationsUniqueId) ? Guid.Parse(item.unreadUserNotificationsUniqueId) : Guid.ECNTy,
                };
            }
            return n;
        }

        private NotificationUserConfigDTO UpdateValues(NotificationUserConfigDTO savedValue, NotificationUserConfig newValues)
        {
            if (savedValue != null)
            {
                savedValue.readedUserNotificationsUniqueId = $"{newValues.hiddenUserNotificationsUniqueId}";
                savedValue.unreadUserNotificationsUniqueId = $"{newValues.unreadUserNotificationsUniqueId}";
            }

            return savedValue;
        }

        #endregion
    }
}
