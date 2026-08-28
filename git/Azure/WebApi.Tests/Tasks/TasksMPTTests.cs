using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Tasks;

public class TasksCNTTests : BasePnPAppTest
{
    // CONFIGURAR .................
    private string BODY_ID = "demoev-cs-voto";
    private string SHARED_EVENT_ID = "ad64f7c2c0a747c6a291e897f12b4148";
    private string UPN = "CNTextconvocante@contoso-test.onmicrosoft.com";

    public TasksCNTTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
    }

    [Fact]
    public async Task GetUserTaskTest()
    {
        try
        {
            var tasksDelegateService = ServiceProvider().GetRequiredService<TasksDelegateService>();
            var taskAttendanceService = ServiceProvider().GetRequiredService<TasksAttendanceService>();
            var taskCertificationService = ServiceProvider().GetRequiredService<TasksCertificationService>();

            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var userTasksprovider = new TasksDALprovider();
            var tasks = await userTasksprovider.GetPendingTaskByEvent(ctx, SHARED_EVENT_ID);
            var taskAttendance = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestAttendance));
            var taskCertification = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestCertificacion));
            var taskDelegation = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestDelegation));
            await taskAttendanceService.CancelOrExpiredAttendance(ctx, BODY_ID, taskAttendance, DALConstants.TaxonomyValuesIds.RequestStatus.Expired);
            await taskCertificationService.CancelOrExpiredCertification(ctx, BODY_ID, taskCertification, DALConstants.TaxonomyValuesIds.RequestStatus.Expired);
            await tasksDelegateService.CancelOrExpiredCertification(ctx, BODY_ID, taskDelegation, DALConstants.TaxonomyValuesIds.RequestStatus.Expired);

            // using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            // var userTasksprovider = new TasksDALprovider();
            // var userToAddOrRemove = await ctx.Web.EnsureUserAsync(UPN);
            // var attendanceUser = await userTasksprovider.GetCancelTaskByIdMeetingAndIdUser(ctx, "1ed317ddb5494b48a67799ea8af12c2b", userToAddOrRemove.Id);
            // var taskAttendanceService = new TasksAttendanceService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksAttendanceService>(), Auth);
            // if (attendanceUser is not null)
            // {
            //     await taskAttendanceService.UpdateCancelTask(ctx, BODY_ID, attendanceUser);
            // }
            // var taskService = new TasksService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth));

            // var listUserTasks = await taskService.GetUserTask(ctx, UPN);

            var result = await Task.Run(() => tasks);

            Assert.True(result.Any());
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    // [Fact]
    // public async Task GetEventsTest()
    // {
    //     try
    //     {
    //         var _simService = new INotificationSmsService(this.GetLogger<INotificationSmsService>(), "", "", "");
    //         var _ProfileService = new ProfileService();
    //         var bodyService = new BodyRoleService(this.MemoryCache, this.PnPContextFactory);
    //         var service = new EventsService(bodyService);
    //         var notificationsService = new NotificationsService(bodyService, _simService, _ProfileService, this.PnPContextFactory, this.MemoryCache);
    //         var notificationBussinessService = new NotificationsBusinessService(Logger as ILogger<NotificationsBusinessService>, bodyService, notificationsService);
    //         var eventCaptureService = new EventCaptureService(Logger as ILogger<EventCaptureService>);
    //         var bodies = await bodyService.GetAllBodiesIds();
    //         var ctx = await CreatePnPContextAsSystem();
    //         var publishedEvents = await service.GetPublishedEventsAsync(ctx,bodies);
    //         var changeEventsInCelebration = publishedEvents.Where(e => e.StartDate > DateTime.UtcNow && e.StartDate < DateTime.UtcNow.AddMinutes(30));
    //         var eventsCelebratedIn24hours = publishedEvents.Where(e => e.StartDate > DateTime.UtcNow && e.StartDate < DateTime.UtcNow.AddDays(1));
    //         var eventsCelebratedIn1hours = publishedEvents.Where(e => e.StartDate > DateTime.UtcNow && e.StartDate < DateTime.UtcNow.AddMinutes(60));
    //         foreach(var e in publishedEvents){
    //         //      await service.ChangeStatusEventBySharedId((IPnPContext)PnPContextFactory, e.BodyId, e.Id, DALConstants.TaxonomyValuesIds.EventStatus.InCelebration);
    //         //      var operation = await eventCaptureService.CaptureAndprepareForSendingEventBySharedId((IPnPContext)PnPContextFactory, e.BodyId, e.Id, "test!!");
    //         //      var instanceId = await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableEventSendingOrchestration), operation, new StartOrchestrationOptions() { InstanceId = $"es-b:{e.BodyId}-e:{e.Id}-c:{operation.ToCaptureState.CaptureId}" });
    //           }

    //         foreach(var e in eventsCelebratedIn24hours){
    //                 SendRememberEventNotification("ma-gr-cyl","32409123kl321jh4ñlk0981234","en-US",RememberNotification.TwentyFourHours);
    //         }

    //         foreach(var e in eventsCelebratedIn1hours){

    //         }

    //         Assert.True(publishedEvents.Any());
    //     }
    //     catch (Exception ex)
    //     {
    //         Assert.Fail(ex.Message);
    //     }
    // }


    [Fact]
    public async Task UpdateTaskById()
    {
        // try
        // {
        //     using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);
        //     var service = new EventsService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<EventsService>(), Auth, ServiceProvider());
        //     var test = await service.GetUserEvents(ctx, BODY_ID);
        //     var result = await Task.Run(() => test);

        //     Assert.True(result.Any());
        // }
        // catch (Exception ex)
        // {
        //     Assert.Fail(ex.Message);
        // }



    }


    [Fact]
    public async Task GetTaskByBodyId()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/demoev-cs-voto/");
            var groups = await ServiceProvider().GetRequiredService<BodyRoleService>().GetBodyGroupsFromUser(BODY_ID, UPN);
            var tasksService = ServiceProvider().GetRequiredService<TasksService>();
            var listUserTasks = await tasksService.GetUserTaskByBodyId(ctx, groups, UPN, BODY_ID); // GetUserTask(ctx, UPN);
            var result = await Task.Run(() => listUserTasks);
            Assert.True(result.Any());
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task GetTaskByTaskId()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/demoev-cs-voto/");
            var tasksService = ServiceProvider().GetRequiredService<TasksService>();
            var listUserTasks = await tasksService.GetUserTaskByTaskId(ctx, BODY_ID, 821);

            //var result = await Task.Run(() => listUserTasks);
            Assert.NotNull(listUserTasks);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }
}
