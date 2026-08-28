
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Bodies;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Domains.Notifications;

public class NotificationsBusinessService(BodyRoleService bodyService, NotificationsService notificationService, ILogger<NotificationsBusinessService> logger, M365AuthHelper auth) : ServiceBasePnP<NotificationsBusinessService>(logger, auth)
{
    private BodyRoleService _bodyService = bodyService;
    private NotificationsService _notificationService = notificationService;

    #region  Métodos Públicos

    public async Task SendRememberEventNotification(IPnPContext ctx, string bodyId, string sharedEventId, RememberNotification reminder)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

        //1. Get Reminder Event
        var eventsDAL = new InConstructionEventsDALprovider();
        var eventDAO = await eventsDAL.GetBySharedEventId(bodyCtx, sharedEventId);
        var remindersDAO = eventDAO.Recordatorios;

        switch (reminder)
        {
            case RememberNotification.OneHour:
                if (!remindersDAO.Contains(NotificationsMessages.Reminders.OneHour))
                {
                    //2. Enviar Recordatorio
                    await SendRememberEventNotification(bodyCtx, bodyId, eventDAO, NotificationsMessages.Meeting.Reminder1H);

                    //3. Guardar Recordatorio
                    remindersDAO.Add(NotificationsMessages.Reminders.OneHour);
                    eventDAO.Recordatorios = remindersDAO;
                    await eventDAO.AsListItem().UpdateAsync();
                }
                break;

            case RememberNotification.TwentyFourHours:
                if (!remindersDAO.Contains(NotificationsMessages.Reminders.TwentyFourHours))
                {
                    //2. Enviar Recordatorio
                    await SendRememberEventNotification(bodyCtx, bodyId, eventDAO, NotificationsMessages.Meeting.Reminder24H);

                    //3. Guardar Recordatorio
                    remindersDAO.Add(NotificationsMessages.Reminders.TwentyFourHours);
                    eventDAO.Recordatorios = remindersDAO;

                    await eventDAO.AsListItem().UpdateAsync();
                }
                break;
        }
    }

    #endregion

    #region  Métodos Privados

    private async Task SendRememberEventNotification(IPnPContext ctx, string bodyId, BaseEventDAO eventDAO, string notificationKey)
    {
        //1. Users
        var usersUPN = eventDAO.Asistentes.AsUserPrincipalName();

        //2. Replace
        var values = new Dictionary<string, string>(){
                {ReplacedTokens.MeetingTitle, eventDAO.Title},
                {ReplacedTokens.DepartmentName, eventDAO.Department?.Label ?? string.ECNTy}
            };

        var url = _bodyService.GetNotificationUrlByEvent(ctx, eventDAO.IdMeeting, bodyId);

        //3. Notification
        await _notificationService.AddUserNotification(bodyId, eventDAO.MeetingType!.TermId, usersUPN, notificationKey, "en-US", url, values);
    }

    #endregion
}