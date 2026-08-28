using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Notifications;

using Contoso.Portal.Domains.Profile;
using Microsoft.Extensions.DependencyInjection;


namespace WebApi.Tests.Events
{
    public class EventPublishingServiceTest : BasePnPAppTest
    {
        // CONFIGURAR .................

        private const string BODY_ID = "ma-gr-cyl";
        private const string SHARED_EVENT_ID = "e7fe8e787f6f4e55b721855c21c52e23";


        private readonly EventPublishingService _eventPublishingService;


        public static TimeSpan Time(Action action)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            action();
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public EventPublishingServiceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
        {
            _eventPublishingService = ServiceProvider().GetRequiredService<EventPublishingService>();
        }

        [Fact]
        public async Task UpdateAttendanceRequestsForEventTest()
        {
            try
            {
                //1. Contexto
                using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);
                var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(BODY_ID))));

                //2. Services
                //var eventsCapture = new EventCaptureService(this.GetLogger<EventCaptureService>(), ---falta PnpContextFactory---);
                //var obj = await eventsCapture.CaptureAndprepareForSendingEventBySharedId(ctx,BODY_ID,SHARED_EVENT_ID, "Cambio Test " + DateTime.Now.ToString("DD/MM/YYYY HH:mm:ss"));

                //await _eventPublishingService.EventCaptureToAttendanceUpdates(bodyCtx,obj);

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task EventCaptureToNotificationsTest()
        {
            try
            {
                //1. Contexto
                using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
                var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(BODY_ID))));

                //2. Services
                var objInformation = new EventSendOperationInformation()
                {
                    BodyId = BODY_ID,

                    //2.1 Estado Meeting                    
                    StatusHasChanged = true,
                    NewEventStatusId = DALConstants.TaxonomyValuesIds.EventStatus.InConstruction,

                    //2.2. Estado Actual
                    CurrentEventStatusId = DALConstants.TaxonomyValuesIds.EventStatus.Published,
                    DetailsHasChanged = true,
                    AttendeesHasChanged = true,
                    DocumentsHasChanged = true,
                };


                await _eventPublishingService.EventCaptureToNotifications(bodyCtx, objInformation);

            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task EventCaptureToTeamsAndOutlookUpdates()
        {
            try
            {
                //1. Contexto
                using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
                var eventsService = ServiceProvider().GetRequiredService<EventsService>();
                var eventPublishingService = ServiceProvider().GetRequiredService<EventPublishingService>();
                var eventCaptureService = ServiceProvider().GetRequiredService<EventCaptureService>();

                // forzar cambio de estado: AQUI SE CREAN los acuerdos y se cambia el estado a "En celebración"
                await eventsService.ChangeStatusByEvent(ctx, BODY_ID, SHARED_EVENT_ID, DALConstants.TaxonomyValuesIds.EventStatus.InCelebration);

                // prodceso de publicación: 
                var result = await eventCaptureService.CaptureAndprepareForSendingByEvent(ctx, BODY_ID, SHARED_EVENT_ID, "Hola hola!!");

                await eventPublishingService.prepareRequirements(ctx, result);
                // await eventPublishingService.EventCaptureToDocumentSetUpdates(ctx, result);
                await eventPublishingService.EventCaptureToTeamsAndOutlookUpdates(ctx, result);
                // await eventPublishingService.EventCaptureToAttendanceUpdates(ctx, result);
                // await eventPublishingService.EventCaptureToNotifications(ctx, result);








                return;



            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

    }
}
