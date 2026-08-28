using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Notifications;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Contoso.Portal.Data.Extensions;
using System.Linq;
using static Contoso.Portal.Data.DAL.DALConstants;
using Microsoft.Extensions.DependencyInjection;

namespace WebApi.Tests.Notifications;

public class NotificationsServiceTests : BasePnPAppTest
{
    // CONFIGURAR .................
    /* private string BODY_ID = "demomvp-gr-pruebados";
     private string SHARED_EVENT_ID = "a41be74f16c345eca6977c47ff261b5e";
     private string USER_UPN = "exttestsidente@contoso.onmicrosoft.com";
     private string URL = "https://test.sharepoint.com/sites/Contoso";*/

    private string BODY_ID = "demomvp-cs-ae";
    private string SHARED_EVENT_ID = "ea44ad7d3a44449d8b91f5399b56506b";
    private string USER_UPN = "gcastrma_emeal.nttdata.com@test.onmicrosoft.com";
    private string URL = "https://test.sharepoint.com/sites/Contoso";

    private readonly NotificationsService _notificationsService;

    public NotificationsServiceTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        _notificationsService = ServiceProvider().GetRequiredService<NotificationsService>();
    }

    [Fact]
    public async Task AddNotificationTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/Contoso");

            var notification = new Notification()
            {
                Locale = "en-US",
                Expires = DateTime.Today.AddDays(2),
                NotificationType = DALConstants.NotificationsMessages.Tipo.Meetings,
                Visualized = false,
            };

            var values = new Dictionary<string, string>();
            values.Add(ReplacedTokens.MeetingTitle, "TESSS");

            await _notificationsService.AddUserNotification(BODY_ID, Guid.ECNTy, USER_UPN, NotificationsMessages.Meeting.Publicar, "en-US", URL, values);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task AddBodyUrgentNotification()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem();

            var notification = new Notification()
            {
                Locale = LocaleContoso.caES,
                Expires = DateTime.Today.AddDays(2),
                NotificationType = DALConstants.NotificationsMessages.Tipo.Sistema,
                Body = "Prueba Comunicación Urgent " + DateTime.Now.ToString("hh:mm:ss"),
                Visualized = false,
            };

            await _notificationsService.UrgentNotification(ctx, BODY_ID, notification, USER_UPN, SHARED_EVENT_ID);
            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetReminderNotificationTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            using var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(BODY_ID))));

            var eventsDAL = new InConstructionEventsDALprovider();
            var eventDAO = await eventsDAL.GetBySharedEventId(bodyCtx, SHARED_EVENT_ID);

            var reminders = eventDAO.Recordatorios.ToList();

            if (!reminders.Contains(NotificationsMessages.Reminders.TwentyFourHours))
            {
                reminders.Add(NotificationsMessages.Reminders.TwentyFourHours);
                eventDAO.Recordatorios = reminders;

                await eventDAO.AsListItem().UpdateAsync();
            }
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.ToString());
        }
    }

    [Fact]
    public async Task GetMessage()
    {
        try
        {
            var values = new Dictionary<string, string>
            {
                {DALConstants.ReplacedTokens.MeetingTitle, "TESSS"}
            };

            var notifications = await _notificationsService.GetMessage(DALConstants.NotificationsMessages.Alert.Urgent
                                                                       , DALConstants."en-US"
                                                                       , values);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetArchivedNotifications()
    {
        try
        {
            // Se crea un contexto
            using var ctx = await CreatePnPContextAsSystem();

            // Se crea un userUPN vacío
            var userUpn = string.ECNTy;

            // Se crea un filtro de fecha para obtener las Notifications archivadas de hace 3 meses
            var dateFilter = DateTime.Now.ToUniversalTime().AddMonths(-3);

            // Se obtienen todas las Notifications archivadas que cumplen con el filtro de fecha
            var expiredArchivedNotifications = await NotificationsService.GetArchivedNotifications(ctx, dateFilter);

            // Se eliminan todas las Notifications archivadas que cumplen con el filtro de fecha
            // await _notificationsService.DeleteArchivedNotifications(ctx, userUpn, expiredArchivedNotifications);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetUserNotifications()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem();

            var notifications = await _notificationsService.GetUserArchivedNotifications(ctx, USER_UPN, "en-US", BODY_ID);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetNotificationById()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem();

            var notification = await _notificationsService.GetNotificationById(ctx, USER_UPN, Guid.Parse("30aaee29-38df-4cf0-be16-9610dae0b74c"));
            Assert.NotNull(notification);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task MarkAsReaded()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem();

            await _notificationsService.UpdateVisualizedUserNotificationById(ctx, "agutiepe@contoso.onmicrosoft.com", Guid.Parse("30aaee29-38df-4cf0-be16-9610dae0b74c"), false); ;

            Assert.True(true);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task MarkAsReadedBatch()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem();

            //"684c00c0-e9f3-49dd-987b-4383240d8bfd"
            var request = new NotificationBatchRequest()
            {
                Operation = NotificationBatchOp.Hide,
                NotificationIds = new List<string>() { "30aaee29-38df-4cf0-be16-9610dae0b74c" }
            };
            await _notificationsService.UpdateNotifications(ctx, "agutiepe@contoso.onmicrosoft.com", request); ;

            Assert.True(true);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

}