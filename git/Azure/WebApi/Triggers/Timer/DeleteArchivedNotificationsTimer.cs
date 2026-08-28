using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Runtime.Helpers;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Timer
{
    public class DeleteArchivedNotificationsTimerHttpTrigger : TriggerBase<DeleteArchivedNotificationsTimerHttpTrigger>
    {
        private NotificationsService _notificationsService;

        // Extestsión cron para activar la función a las 00:00:00 del primer domingo de cada mes
        private const string CronExtestssion = "0 0 0 1-7 * 0";

        /* Constructor */
        public DeleteArchivedNotificationsTimerHttpTrigger(NotificationsService notificationsService, ILogger<DeleteArchivedNotificationsTimerHttpTrigger> logger, M365AuthHelper auth) : base(logger, auth)
        {
            this._notificationsService = notificationsService;
        }

        /* Método que será llamado cada vez que se active el temporizador */
        [Function(nameof(DeleteArchivedNotifications))]
        public async Task DeleteArchivedNotifications([TimerTrigger(CronExtestssion)] FunctionContext context)
        {
            try
            {
                // Se crea un contexto
                using var ctx = await CreatePnPContextAsSystem();

                // Se crea un filtro de fecha para obtener las Notifications archivadas de hace 6 meses
                var dateFilter = DateTime.Now.ToUniversalTime().AddDays(-180);

                // Se obtienen todas las Notifications archivadas que cumplen con el filtro de fecha
                var expiredArchivedNotifications = await NotificationsService.GetArchivedNotifications(ctx, dateFilter);

                // Se eliminan todas las Notifications archivadas que cumplen con el filtro de fecha
                await NotificationsService.DeleteArchivedNotifications(expiredArchivedNotifications);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error DeleteArchivedNotifications", ex);
            }
        }
    }
}