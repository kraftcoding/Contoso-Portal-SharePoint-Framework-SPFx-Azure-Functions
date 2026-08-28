using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Bodies;

public class BodyConfigTest : BasePnPAppTest
{
    // CONFIGURAR .................
    private string BODY_ID = "ma-gr-cyl";

    private string LOCALE = "es-ES";

    private readonly string taxonomySiteId = "";
    private readonly string appTermGroup = "";
    private readonly string setId = "";



    private readonly ConfigDepartmentsService _ConfigDepartmentsService;

    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public BodyConfigTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {

        // _ConfigDepartmentsService = new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth);
        _ConfigDepartmentsService = new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth, LOCALE, taxonomySiteId, appTermGroup, setId);
    }


    [Fact]
    public async Task GetBodiesEmail()
    {
        try
        {
            await _ConfigDepartmentsService.GetBodiesEmail(BODY_ID, Guid.ECNTy, DALConstants.EmailBodyTypes.ReservadaAgenda, LOCALE);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }


    [Fact]
    public async Task GetPeriodDaysApprodvalMinutesTest()
    {
        try
        {

            var result = await _ConfigDepartmentsService.GetPeriodDaysApprodvalMinutes(BODY_ID);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }
}