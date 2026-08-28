using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Tasks;

public class TasksApprodvalMinutesTest : BasePnPAppTest
{
    // CONFIGURAR .................

    private string BODY_ID = "ma-gr-cyl";
    private string SHARED_EVENT_ID = "7e847c7e62f345f8a6806b32b192e11e";


    public TasksApprodvalMinutesTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {


    }

    [Fact]
    public async Task AddTaskTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksApprovalMinutesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksApprovalMinutesService>(), Auth, ServiceProvider());
            var deltask = new TasksApprodvalMinutes()
            {
                //Requeridos
                SharedEventId = SHARED_EVENT_ID,
                RequestedBy = "admin@contoso.onmicrosoft.com",
                AssignedTo = new List<string> { "fnietolo@contoso.onmicrosoft.com" },
                IdMinutes = "",

                //Opcionales
                TaskStartDate = DateTime.Now,
                TaskEndDate = DateTime.Now.AddHours(1).AddMinutes(1),
                Comment = "Test Comentario 123 .,|$%&/()-_",
            };

            await service.AddTask(ctx, BODY_ID, deltask);

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
            var service = new TasksApprovalMinutesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksApprovalMinutesService>(), Auth, ServiceProvider());

            var task = new TasksApprodvalMinutes()
            {
                TaskId = "324",
                TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.PendingModification,
                Comment = "Modificado",
                TaskStartDate = DateTime.Now,
            };

            await service.UpdateTask(ctx, BODY_ID, task);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task GetTaskStatusApprodvalMinutesTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

            var service = new TasksApprovalMinutesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksApprovalMinutesService>(), Auth, ServiceProvider());

            await service.GetStatusApprodvalMinutes(ctx, BODY_ID, SHARED_EVENT_ID);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }






}