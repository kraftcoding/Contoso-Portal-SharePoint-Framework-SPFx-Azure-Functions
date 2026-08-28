using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Graph;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Tasks;

public class TasksModificationMinutesTest : BasePnPAppTest
{
    // CONFIGURAR .................
    private string BODY_ID = "ma-gr-cyl";
    private string SHARED_EVENT_ID = "2736cf55a16145b2b8371ccd98d2d941";


    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public TasksModificationMinutesTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task AddTaskTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksModificationMinutesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksModificationMinutesService>(), Auth, Serviceprovider());

            //SolicitadoPor, es el AsignadoA en la tarea padre.
            var task = new TasksModificationMinutes()
            {
                //Requeridos
                SharedEventId = SHARED_EVENT_ID,
                RequestedBy = "fnieto@contoso.onmicrosoft.com",
                TaskParentId = "316",
                TaskTitle = "Test Convo",

                //Opcionales
                TaskStartDate = DateTime.Now,
                TaskEndDate = DateTime.Now.AddHours(1).AddMinutes(1),
                Comment = "Test Comentario 123 .,|$%&/()-_",
            };

            await service.AddTask(ctx, BODY_ID, task);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task UpdateTaskTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksModificationMinutesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksModificationMinutesService>(), Auth, Serviceprovider());

            //SolicitadoPor, es el AsignadoA en la tarea padre.
            var deltask = new TasksModificationMinutes()
            {
                TaskId = "585",
                IdMinutes = "2135131351351",
                TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Accepted,
            };

            await service.UpdateTask(ctx, BODY_ID, deltask);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

}