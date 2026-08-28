using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Contoso.Portal.Domains.Bodies;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests;

public class BodyByIdTest : BasePnPAppTest
{
    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public BodyByIdTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task GetBodyById()
    {
        using var ctx = await CreatePnPContextAsUser("sites/contoso-dept-main/");
        var service = new BodiesService(new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth), Log<BodiesService>(), Auth);
        //var result = await service.GetBodyById(ctx, "contoso-dept-main");
    }
}