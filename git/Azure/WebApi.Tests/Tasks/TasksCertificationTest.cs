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

public class TasksCertificationTest : BasePnPAppTest
{
    // CONFIGURAR .................

    private string BODY_ID = "demomvp-cs-ae";
    private string SHARED_EVENT_ID = "7e847c7e62f345f8a6806b32b192e11e";


    public TasksCertificationTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {


    }

    [Fact]
    public async Task AddTaskTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new TasksCertificationService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksCertificationService>(), Auth);

            var task = new TasksCertification()
            {
                //Requeridos
                SharedEventId = SHARED_EVENT_ID,
                RequestedBy = "admin@contoso.onmicrosoft.com",
                AgreementSharedId = "1",

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
            var service = new TasksCertificationService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<TasksCertificationService>(), Auth);

            var task = new TasksCertification()
            {
                TaskId = "490",
                Comment = "Modificado",
                TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Accepted,
            };

            await service.UpdateTask(ctx, BODY_ID, task);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }
}