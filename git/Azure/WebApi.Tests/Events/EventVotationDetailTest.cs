using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Model.Events;
using PnP.Core.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace WebApi.Tests.Events
{
    public class EventVotationDetailTest : BasePnPAppTest
    {
        // CONFIGURAR .................

        private string BODY_ID = "demoev-cs-voto";
        private string SHARED_EVENT_ID = "8defff3315524370a795d6a4c3600315";

        private EventsService _eventsService;

        private EventAgendaItemsService _eventAgendaItemService;
        private EventAgreementService _eventAgreementService;
        private EventCaptureService _eventCaptureService;

        private EventVotationService _eventVotationService;

        public static TimeSpan Time(Action action)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            action();
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }


        public EventVotationDetailTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
        {
            _eventsService = Serviceprovider().GetRequiredService<EventsService>();
            _eventCaptureService = Serviceprovider().GetRequiredService<EventCaptureService>();
            _eventAgendaItemService = Serviceprovider().GetRequiredService<EventAgendaItemsService>();
            _eventAgreementService = Serviceprovider().GetRequiredService<EventAgreementService>();
            _eventVotationService = Serviceprovider().GetRequiredService<EventVotationService>();
        }

        [Fact]
        public async Task GetVotersRetestsentation()
        {
            try
            {
                var ctx = await CreatePnPContextAsSystem("/sites/" + BODY_ID);
                var voters = await _eventVotationService.GetRetestsentationsByUser(ctx, BODY_ID, SHARED_EVENT_ID);
                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task StartVoters()
        {
            try
            {
                BODY_ID = "demomvp-cs-ae";
                SHARED_EVENT_ID = "b3a4992c321b46e492222d7e1c35a874";
                var ctx = await CreatePnPContextAsSystem("/sites/" + BODY_ID);
                var voters = await _eventVotationService.StartVotation(ctx, BODY_ID, SHARED_EVENT_ID);
                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task CreatevOTATION()
        {
            try
            {
                var ctx = await CreatePnPContextAsSystem("/sites/" + BODY_ID);
                var votationprovider = new EventVotationDetailDALprovider(new InConstructionEventsDALprovider());
                var eventVotation = new EventVotationDetailDTO()
                {
                    ComunidadAsignada = VoteRetestsentations.Scheduler,
                    AsignadoA = []
                };
                await votationprovider.AddOrUpdateItemsForEvent(ctx, SHARED_EVENT_ID, [eventVotation]);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task UpdateVote()
        {
            try
            {
                BODY_ID = "demomvp-cs-ae";
                SHARED_EVENT_ID = "deec995fb46d4c1d93f027ad4f1ddd63";
                var ctx = await CreatePnPContextAsUser("/sites/" + BODY_ID);
                var votationprovider = new EventVotationDetailDALprovider(new InConstructionEventsDALprovider());
                var eventVotation = new Vote()
                {
                    Id = "0e020a919ee84cf682e0b0af7714d95e",
                    VoteId = "bd558d1a-e24b-4882-865d-853c7a14289e"
                };
                await _eventVotationService.UpdateVoteByUser(ctx, BODY_ID, SHARED_EVENT_ID, eventVotation.Id, eventVotation.VoteId, "CNTextasistentemiembro@contoso-test.onmicrosoft.com");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task UpdateVoteByRetestsentation()
        {
            try
            {
                var ctx = await CreatePnPContextAsSystem("/sites/" + BODY_ID);
                var votationprovider = new EventVotationDetailDALprovider(new InConstructionEventsDALprovider());
                var vote = new Dictionary<string, string>
                {
                    { "c8ba5a5a41bd4e7ebce0fdf2ed4d5846", "0ae772d2-0f28-4eaf-95a9-8c857d78d51c" }
                };
                var votation = new EventVotation()
                {
                    UniqueSharedID = "0471dbaf935e422bab8747a85d24e534",
                    VoteRetestsentationId = "5db8575a-b9ad-44ae-9e7f-7e95add703af",
                    AssignedUsers = [],
                    Votes = vote

                };

                var votationRes = await _eventVotationService.AddOrUpdateVotationForEvent(ctx, BODY_ID, SHARED_EVENT_ID, [votation]);
                Assert.True(true);

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task CheckModifyAgreementsByVotationStart()
        {
            try
            {
                var ctx = await CreatePnPContextAsSystem("/sites/" + BODY_ID);

                await _eventVotationService.AgreementStatusToStartVotationSession(ctx, BODY_ID, SHARED_EVENT_ID);
                Assert.True(true);

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }
    }


}
