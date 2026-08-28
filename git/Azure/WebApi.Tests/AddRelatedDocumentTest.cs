using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests;

public class AddRelatedDocumentTest : BasePnPAppTest
{
    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public AddRelatedDocumentTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task GetBodyById()
    {

        // try
        // {
        //     using var ctx = await CreatePnPContextAsUser("sites/ma-gr-cyl/");
        //     var service = new EventsService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<EventsService>(), Auth, Serviceprovider());
        //     var bodyId = "ma-gr-cyl";
        //     var eventId = "db1e5b9b3a3e45359d58c9a3aedaacbe";
        //     var idAgreement = "433";
        //     await service.RemoveRelatedDocumentIdToAgreement(ctx, bodyId, eventId, idAgreement, "gdfgdfgdfg");
        // }
        // catch (Exception ex)
        // {
        //     Assert.Fail(ex.Message);
        // }

    }
}