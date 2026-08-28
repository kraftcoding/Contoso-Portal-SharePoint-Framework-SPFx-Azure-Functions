using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Events;
using Newtonsoft.Json;
using PnP.Core.Model.Security;
using PnP.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Attendance;

public class AttendanceTest : BasePnPAppTest
{
    // CONFIGURAR .................

    private string BODY_ID = "demomvp-cs-aedos";
    private string SHARED_EVENT_ID = "00fbdf6cebcf4c8fae7a4e658fce2544";


    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public AttendanceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task GetAttendaceByBody()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var service = new EventAttendanceService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<EventAttendanceService>(), Auth, Serviceprovider());

            var results = await service.GetEventAttendance(ctx, BODY_ID, SHARED_EVENT_ID);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }


    [Fact]
    public async Task DeleteAttendance()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/ma-gr-cyl/");
            var service = new EventAttendanceService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<EventAttendanceService>(), Auth, Serviceprovider());
            //  string bodyId = "ma-cs-adr";
            //  string eventId = "200";
            string bodyId = "ma-gr-cyl";
            string eventId = "728eae70143c44e68d0f522fb829fd51";
            string upn = "fnieto@contoso.onmicrosoft.com";
            bool isAdd = false;
            // var userTasksprovider = new TasksDALprovider();
            // var delegatedTasks = await userTasksprovider.GetDelegatedTaskByIdMeetingAndIdUserDelegated(ctx, eventId, 11);
            await service.AddOrRemoveEventAttendee(ctx, bodyId, eventId, upn, isAdd);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }


    }

    [Fact]
    public async Task UpdateCheckAttendance()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/aa-co-ccc/");
            var service = new EventAttendanceService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<EventAttendanceService>(), Auth, Serviceprovider());
            string bodyId = "aa-co-ccc";
            string eventId = "0098b3b15e604212900799c59732c902";
            string jsonAttendance = "[{\"UserPrincipalName\": \"agutiepe@contoso.onmicrosoft.com\",\"RequestStatusId\": \"b6d26a81-48bb-4a65-a0d6-7d50b3feed1a\",\"RequestUserMessage\": \"\",\"RequestStatusUpdate\": \"0001-01-01T00:00:00\",\"HasAttended\": false,\"AttendanceTypeId\": \"474ce74e-26c2-4ada-9dbb-516be5deb394\",\"DelegationUserPrincipalName\": \"\",\"Vote\": false,\"IsGuest\": false,\"AttendanceId\": \"144\",\"AttendanceParentId\": \"\"},{\"UserPrincipalName\": \"falonspa@contoso.onmicrosoft.com\",\"RequestStatusId\": \"b6d26a81-48bb-4a65-a0d6-7d50b3feed1a\",\"RequestUserMessage\": \"\",\"RequestStatusUpdate\": \"0001-01-01T00:00:00\",\"HasAttended\": true,\"AttendanceTypeId\": \"474ce74e-26c2-4ada-9dbb-516be5deb394\",\"DelegationUserPrincipalName\": \"\",\"Vote\": false,\"IsGuest\": false,\"AttendanceId\": \"143\",\"AttendanceParentId\": \"\"}]";

            EventUserAttendance[] eventUserAttendance = JsonConvert.DeserializeObject<EventUserAttendance[]>(jsonAttendance);

            var results = await service.UpdateAttendance(ctx, bodyId, eventId, eventUserAttendance.ToList());
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task DownloadAttendanceTemplateTest()
    {
        var eventAttendanceTemplateService = Serviceprovider().GetRequiredService<EventAttendanceTemplateService>();
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + "demomvp-cs-aedos"); //using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);
            var response = await eventAttendanceTemplateService.DownloadAttendaceTemplate(ctx, "demomvp-cs-aedos", "dd1bb2d35d8248b6b930b0d9e2f2fd92");

            string ficheroGuarda = @"./Document.pdf";
            if (System.IO.File.Exists(ficheroGuarda)) try { System.IO.File.Delete(ficheroGuarda); } catch (Exception) { }

            string whereSave = System.IO.Directory.GetCurrentDirectory();
            Debug.WriteLine("PRUEBA::: localización del fichero generado= " + whereSave);
            File.WriteAllBytes(ficheroGuarda, response);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }
}