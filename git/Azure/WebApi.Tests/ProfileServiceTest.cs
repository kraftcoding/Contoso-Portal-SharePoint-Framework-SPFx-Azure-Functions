using Contoso.Portal.Domains.profile;
using Contoso.Portal.Model.profile;
using System;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests;

public class profileServiceTests : BasePnPAppTest
{
    private string SiteTest = "sites/demomvp-cs-aedos";
    public profileServiceTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task UpdateprofileTest()
    {
        try
        {
            Random rand = new Random();
            //using var ctx = await CreatePnPContextAsSystem(SiteTest);
            using var ctx = await CreatePnPContextAsUser(SiteTest);

            var profileService = new profileService();
            var updateprofile = await profileService.Getprofile(ctx, "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com");
            updateprofile.CellPhone = "777777777";
            await profileService.Updateprofile(ctx, "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com", updateprofile);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetprofileTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsUser(SiteTest);
            var profileService = new profileService();

            var updateprofile = await profileService.Getprofile(ctx, "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com");

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

}
