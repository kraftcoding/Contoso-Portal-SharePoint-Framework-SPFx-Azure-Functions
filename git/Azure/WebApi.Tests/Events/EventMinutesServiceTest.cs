using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Model.Tasks;
using Xunit;
using Xunit.Abstractions;
using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;

namespace WebApi.Tests.Events;

public class EventMinutesServiceTest : BasePnPAppTest
{
    private string BODY_ID = "demomvp-cs-aedos";
    private string SHARED_EVENT_ID = "00fbdf6cebcf4c8fae7a4e658fce2544";

    private string DOC_ID = "3dfe5a68-7f6e-47ec-af15-fab34f3ffb8c";
    private readonly EventMinutesService service;


    public EventMinutesServiceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        var bodyRoleService = new BodyRoleService(MemoryCache, Log<BodyRoleService>(), Auth);
        service = Serviceprovider().GetRequiredService<EventMinutesService>();
    }


    [Fact]
    public async Task GetMinutesInformationTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var result = await service.GetMinutesInformationByEvent(ctx, BODY_ID, SHARED_EVENT_ID);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task SetMinutesInformationTest()
    {
        try
        {

            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

            var minutesInfo = new EventMinutesInformation()
            {
                AgreementsInformation = "InfAcuerdos 123 .,|$%&/()-_",
                AttendanceInformation = "InfAttendance 123 .,|$%&/()-_",
                DocumentationInformation = "InfDocumentacion 123 .,|$%&/()-_",
                AgendaInformation = "InfAgendaItem 123 .,|$%&/()-_",
            };

            await service.SetMinutesInformation(ctx, BODY_ID, SHARED_EVENT_ID, minutesInfo);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task SetMinutesDocumentTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

            await service.SetMinutesDocument(ctx, BODY_ID, SHARED_EVENT_ID, DOC_ID);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

}