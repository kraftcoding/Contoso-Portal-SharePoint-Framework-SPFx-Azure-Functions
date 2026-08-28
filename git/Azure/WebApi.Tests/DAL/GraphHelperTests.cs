using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.DAL
{
    public class GraphHelperTests : BasePnPAppTest
    {
        private GraphHelper _graphClient;
        private IPnPContext _ctx;
        public GraphHelperTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

        private async Task Init(string url = null)
        {
            _ctx = string.IsNullOrWhiteSpace(url) ? await CreatePnPContextAsSystem() : await CreatePnPContextAsSystem(url);
            _graphClient = new GraphHelper(_ctx);
        }

        [Fact]
        public async Task GetUserIdTest()
        {
            try
            {
                await Init();
                var groups = await _graphClient.GetUserGroups("a471221a-f0b0-40e8-b8b8-5d22f6f217bb");
                Assert.NotNull(groups);
                Assert.NotECNTy(groups);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task GetGroupsMembersStartsWith()
        {
            try
            {
                await Init();
                var groups = await _graphClient.GetGroupsMembersWhenNameStartsWith("demoev-cs-voto");
                Assert.NotNull(groups);
                Assert.NotECNTy(groups);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }


        [Fact]
        public async Task GetprofileBatch()
        {
            try
            {
                await Init();
                var upns = new List<string>()
                {
                    "agutiepe@contoso.onmicrosoft.com",
                    "acallejo@contoso.onmicrosoft.com",
                    "azize@contoso.onmicrosoft.com",
                    "cgronset@contoso.onmicrosoft.com",
                    "admin@contoso.onmicrosoft.com",
                    "ep1091_outlook.com@contoso.onmicrosoft.com",
                    "eduapaul@contoso.onmicrosoft.com",
                    "ntteduapaul@contoso.onmicrosoft.com",
                    "fnieto@contoso.onmicrosoft.com",
                    "fernando.alonso.pacheco_emeal.nttdata.com#EXT#@contoso.onmicrosoft.com",
                    "falonspa@contoso.onmicrosoft.com",
                    "fernando.david.nieto.lopez_emeal.nttdata.com#EXT#@contoso.onmicrosoft.com",
                    "javier@contoso.onmicrosoft.com",
                    "javier.lopez.sanchez_emeal.nttdata.com#EXT#@contoso.onmicrosoft.com",
                    "noreplay@contoso.onmicrosoft.com",
                    "portegal@contoso.onmicrosoft.com",
                    "Resource1@contoso.onmicrosoft.com",
                    "room1@contoso.onmicrosoft.com",
                    "Room2@contoso.onmicrosoft.com",
                    "nolicenciauser@contoso.onmicrosoft.com",
                    "victor@contoso.onmicrosoft.com"
                };
                var profiles = await _graphClient.GetprofileInBatch(upns, new List<string>() { DALConstants.Userproperties.Graph.Mail, DALConstants.Userproperties.Graph.OtherMails });
                Assert.NotNull(profiles);
                Assert.NotECNTy(profiles);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task CaptureDocumentSet()
        {
            try
            {
                await Init("sites/ma-cs-adr/");
                await _ctx.Site.EnsurepropertiesAsync(s => s.Id);
                var captureId = await _graphClient.CaptureDocumentSet(_ctx.Site.Id.ToString(), "8cd33d2f-af9a-45da-a0a1-379f5b82e2aa", "389", false, "Testing back");
                Assert.True(captureId > 0);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task GetChangesFromDocumentSetVersion()
        {
            try
            {
                await Init("sites/ma-cs-adr/");
                await _ctx.Site.EnsurepropertiesAsync(s => s.Id);
                var diff = await _graphClient.GetChangesFromDocumentSetVersion(_ctx.Site.Id.ToString(), "8cd33d2f-af9a-45da-a0a1-379f5b82e2aa", "389", "1", "2");
                Assert.NotNull(diff);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task DownloadSpecificVersionOfFile()
        {
            try
            {
                await Init("sites/ma-cs-adr/");
                await _ctx.Site.EnsurepropertiesAsync(s => s.Id);
                var filePath = await _graphClient.DownloadSpecificVersionOfFile(_ctx.GraphClient.Client, _ctx.Site.Id.ToString(), "8cd33d2f-af9a-45da-a0a1-379f5b82e2aa", "a156baef-42e2-4e84-a3a4-a6082219f5ef", "1.0", "MyTemplate2.docx");
                Assert.NotNull(filePath);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task CreateMeetingEvent()
        {
            try
            {
                await Init("sites/ma-cs-adr/");
                await _ctx.Site.EnsurepropertiesAsync(s => s.Id);
                var createdEvent = await _graphClient.CreateMeetingEvent("noreplay@contoso.onmicrosoft.com", new Microsoft.Graph.Models.Event()
                {
                    ResponseRequested = false,
                    AllowNewTimeprodposals = false,
                    Subject = "Test back",
                    Start = new Microsoft.Graph.Models.DateTimeTimeZone()
                    {
                        DateTime = (DateTime.Today).ToString("yyyy-MM-ddTHH:mm:ss"),
                        TimeZone = "UTC"
                    },
                    End = new Microsoft.Graph.Models.DateTimeTimeZone()
                    {
                        DateTime = DateTime.Today.AddDays(2).ToString("yyyy-MM-ddTHH:mm:ss"),
                        TimeZone = "UTC"
                    }
                });

                Assert.NotNull(createdEvent);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        [Fact]
        public async Task GetMeetingEventById()
        {
            try
            {
                await Init("sites/ma-cs-adr/");
                await _ctx.Site.EnsurepropertiesAsync(s => s.Id);
                var createdEvent = await _graphClient.GetMeetingEventById("a9d64fdf-85c9-49f0-8cf7-caefd75d0dbb", "AAMkAGQ4OGMyMDRlLTNkNWUtNDAzZC1hZmU3LWRhYzJjMDEzZjhmNQBGAAAAAABvgMRGSbFsTINaN9kEa_wLBwAihVBdxEnOSITTvsqpchilAAAAAAENAAAihVBdxEnOSITTvsqpchilAAAH5AcJAAA=");
                Assert.NotNull(createdEvent);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }
    }
}
