using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Tasks;
using Contoso.Portal.Triggers.Http.EventsHttpTrigger;
using PnP.Core.Services;
using Xunit;
using Xunit.Abstractions;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace WebApi.Tests.Timer;

public class TimerTest : BasePnPAppTest
{

    private readonly NotificationsService _notificationsService;
    private readonly NotificationsBusinessService _notificationsBussinessService;
    private readonly BodyRoleService _bodyRoleService;
    private readonly EventCaptureService _eventCaptureService;
    private EventsService _eventService;
    private ConfigDepartmentsService _configBodyService;
    private TasksApprovalMinutesService _taskMinutesService;
    private EventMinutesService _eventMinutesService;

    private string BODY_ID = "ma-gr-cyl";
    private string SHARED_EVENT_ID = "90787880fa1b44a5a9f96260ceae3901";
    private string USER_UPN = "javier@contoso.onmicrosoft.com";
    private string URL = "https://contoso-dev.sharepoint.com/sites/Contoso";

    public TimerTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        _bodyRoleService = ServiceProvider().GetRequiredService<BodyRoleService>();
        _eventService = ServiceProvider().GetRequiredService<EventsService>();
        _notificationsService = ServiceProvider().GetRequiredService<NotificationsService>();
        _notificationsBussinessService = ServiceProvider().GetRequiredService<NotificationsBusinessService>();
        _eventCaptureService = ServiceProvider().GetRequiredService<EventCaptureService>();
        _configBodyService = ServiceProvider().GetRequiredService<ConfigDepartmentsService>();
        _taskMinutesService = ServiceProvider().GetRequiredService<TasksApprovalMinutesService>();
        _eventMinutesService = ServiceProvider().GetRequiredService<EventMinutesService>();
    }

    public DurableTaskClient client { get; }

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
    public async Task DeleteArchivedNotifications()
    {
        try
        {
            // Se crea un contexto
            using var ctx = await CreatePnPContextAsSystem();

            // Se crea un filtro de fecha para obtener las Notifications archivadas de hace 6 meses
            var dateFilter = DateTime.Now.ToUniversalTime().AddDays(-79);

            // Se obtienen todas las Notifications archivadas que cumplen con el filtro de fecha
            var expiredArchivedNotifications = await NotificationsService.GetArchivedNotifications(ctx, dateFilter);

            // Se eliminan todas las Notifications archivadas que cumplen con el filtro de fecha
            await NotificationsService.DeleteArchivedNotifications(expiredArchivedNotifications);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }
}