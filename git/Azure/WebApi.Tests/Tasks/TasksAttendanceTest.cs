using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Graph;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Tasks;

public class TasksAttendanceTest : BasePnPAppTest
{
    // CONFIGURAR .................

    private string BODY_ID = "ma-gr-cyl";
    private string SHARED_EVENT_ID = "e90dcea203464636a6c5d41af6e29612";

    // private string BODY_ID ="ma-cs-adr";
    //private string SHARED_EVENT_ID ="c9a27d2d5eb54cb398be9206b047cd58";

    public TasksAttendanceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task AddTaskTest()
    {
        try
        {
            // using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            // var service = new TasksAttendanceService(Log<TasksAttendanceService>(), Auth);

            // var deltask = new TasksAttendance()
            // {
            //     //Requeridos
            //     SharedEventId = SHARED_EVENT_ID,
            //     RequestedBy = "admin@contoso.onmicrosoft.com",
            //     AssignedTo = new List<string> { "agutiepe@contoso.onmicrosoft.com" },

            //     //Opcionales
            //     TasksAttendanceTypeId = DALConstants.TaxonomyValuesIds.AttendanceType.InPerson,
            //     TaskStartDate = DateTime.Now,
            //     TaskEndDate = DateTime.Now.AddHours(1).AddMinutes(1),
            //     Vote = true,
            //     HasAttended = true,
            //     Comment = "Test Comentario 123 .,|$%&/()-_",
            // };

            // await service.AddTask(ctx, BODY_ID, deltask);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task DelegateTest()
    {
        try
        {
            //     using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            //     var service = new TasksAttendanceService(Log<TasksAttendanceService>(), Auth);
            //     var deltask = new TasksAttendance()
            //     {

            //         TaskId = "113",
            //         SharedEventId = "18",
            //         Comment = "Modificado " + DateTime.Now.ToString("hh_mm_ss"),
            //     };

            //     await service.AddTask(ctx, BODY_ID, deltask);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateTask()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = Serviceprovider().GetRequiredService<TasksAttendanceService>();



            //var results = await service.UpdateTaskItem(ctx, "contoso-dept-main","20", termGuid , message);

            // var deltask = new TasksAttendance(){      
            //     //Requeridos
            //     TaskId="113",
            //     SharedEventId = SHARED_EVENT_ID,
            //     RequestedBy= "admin@contoso.onmicrosoft.com",
            //     AssignedTo = new List<string>{"agutiepe@contoso.onmicrosoft.com"},                    

            //     Comment = "Tarea Delegada " + DateTime.Now.ToString("hh_mm_ss"),    

            // };       

            // var task = new TasksAttendance()
            // {
            //     TaskId = "23",
            //     TaskStatusId = "9628620a-c31b-4903-9bb2-b305d51821bd",
            //     Comment = "test",
            // };

            // await service.AddTask(ctx, BODY_ID, task);



        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }






}