using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace WebApi.Tests.Tasks;

public class TasksDelegateTest : BasePnPAppTest
{
    // CONFIGURAR .................
    private string BODY_ID = "demoev-cs-voto";
    private string SHARED_EVENT_ID = "7faee317e1fc40999c644df2b758b4fb";


    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public TasksDelegateTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    // [Fact]
    // public async Task AddTaskTest()
    // {
    //     try
    //     {
    //         using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
    //         var service = new TasksDelegateService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksDelegateService>(), Auth, ServiceProvider());

    //         //SolicitadoPor, es el AsignadoA en la tarea padre.
    //         var deltask = new TasksDelegate()
    //         {
    //             //Requeridos
    //             SharedEventId = SHARED_EVENT_ID,
    //             RequestedBy = "javier@contoso.onmicrosoft.com",
    //             DelegateTo = "victor@contoso.onmicrosoft.com",
    //             AssignedTo = new List<string> { "fnieto@contoso.onmicrosoft.com" },
    //             DelegateVote = true,
    //             TaskParentId = "721",
    //             TaskTitle = "Convo Test 2211",

    //             //Opcionales
    //             TaskStartDate = DateTime.Now,
    //             TaskEndDate = DateTime.Now.AddHours(1).AddMinutes(1),
    //             Comment = "Test Comentario 123 .,|$%&/()-_",
    //         };

    //         await service.AddTask(ctx, BODY_ID, deltask);

    //     }
    //     catch (Exception ex)
    //     {
    //         Assert.Fail(ex.Message);
    //     }

    // }

    [Fact]
    public async Task UpdateTaskTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksDelegateService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksDelegateService>(), Auth, ServiceProvider(), ServiceProvider().GetRequiredService<NotificationsService>(), ServiceProvider().GetRequiredService<TasksAttendanceService>(), DALConstants."en-US");

            //SolicitadoPor, es el AsignadoA en la tarea padre.
            var deltask = new TasksDelegate()
            { RequestedBy = "gcastrma_emeal.nttdata.com@contoso-test.onmicrosoft.com", TaskTitle = "Prueba Meeting nuevo template", TaskStartDate = DateTime.Parse("2024-08-12T11:52:01+00:00"), TaskEndDate = DateTime.Parse("2024-07-26T10:31:43+00:00"), Comment = "", SharedEventId = "39d7e0a16f354ba4a63a692ea8155327", TaskParentId = "264", DelegateTo = "CNTextmiembro1@contoso-test.onmicrosoft.com", DelegateVote = true, DelegatedUserVote = "CNTextmiembro1@contoso-test.onmicrosoft.com" };

            await service.AddTask(ctx, BODY_ID, deltask);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task NewDelegationApprodvalTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksDelegateService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksDelegateService>(), Auth, ServiceProvider(), ServiceProvider().GetRequiredService<NotificationsService>(), ServiceProvider().GetRequiredService<TasksAttendanceService>(), DALConstants."en-US");

            //SolicitadoPor, es el AsignadoA en la tarea padre.
            var deltask = new TasksDelegate()
            {
                SharedEventId = SHARED_EVENT_ID,
                TaskId = "1977",
                TaskStatusId = TaxonomyValuesIds.RequestStatus.Accepted,
            };

            await service.UpdateTask(ctx, BODY_ID, deltask);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

}