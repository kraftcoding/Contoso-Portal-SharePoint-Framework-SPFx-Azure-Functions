using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
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

namespace WebApi.Tests.Events
{
    public class EventTest : BasePnPAppTest
    {
        // CONFIGURAR .................

        private string BODY_ID = "demoev-cs-voto";
        private string SHARED_EVENT_ID = "4ca305d6cef34a3899ff01001b8201cc";

        private EventsService _eventsService;

        private EventAgendaItemsService _eventAgendaItemService;
        private EventAgreementService _eventAgreementService;
        private readonly EventCaptureService _eventCaptureService;

        public static TimeSpan Time(Action action)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            action();
            stopwatch.Stop();
            return stopwatch.Elapsed;
        }


        public EventTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
        {
            _eventsService = Serviceprovider().GetRequiredService<EventsService>();
            _eventCaptureService = Serviceprovider().GetRequiredService<EventCaptureService>();
            _eventAgendaItemService = Serviceprovider().GetRequiredService<EventAgendaItemsService>();
            _eventAgreementService = Serviceprovider().GetRequiredService<EventAgreementService>();
        }

        [Fact]
        public async Task CreateEventVotationDetails()
        {
            try
            {
                // using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");
                // var items = new List<EventAgendaItem>();
                // var item1 =
                //     new EventAgendaItem()
                //     {
                //         Id = 0,
                //         Title = "Prueba item decisorio 1",
                //         Description = "Prueba item 1",
                //         Duration = 10.0,
                //         Order = 1,
                //         AgendaItemType = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5",
                //         RelatedDocumentsIds = [],
                //         OrderType = "c2c12ed4-1001-466f-977c-07249f45c3a0",
                //         ParentId = ""
                //     };

                // // var item2 =
                // //  new EventAgendaItem()
                // //  {
                // //      Id = 0,
                // //      Title = "Prueba item decisorio",
                // //      Description = "Prueba item 2",
                // //      Duration = 10.0,
                // //      Order = 2,
                // //      AgendaItemType = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5",
                // //      RelatedDocumentsIds = [],
                // //      OrderType = "c2c12ed4-1001-466f-977c-07249f45c3a0",
                // //      ParentId = ""
                // //  };
                // items.Add(item1);
                // // items.Add(item2);
                // var results = await _eventAgendaItemService.AddOrUpdateAgendaItemsForEvent(ctx, BODY_ID, SHARED_EVENT_ID, items.ToArray());


                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task AddAgendaItem()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");
                var items = new List<EventAgendaItem>();
                var item1 =
                    new EventAgendaItem()
                    {
                        Id = "",
                        Title = "Prueba item decisorio 1",
                        Description = "Prueba item 1",
                        Duration = 10.0,
                        Order = 1,
                        AgendaItemType = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5",
                        RelatedDocumentsIds = [],
                        OrderType = "c2c12ed4-1001-466f-977c-07249f45c3a0",
                        ParentId = ""
                    };

                var item2 =
                 new EventAgendaItem()
                 {
                     Id = "",
                     Title = "Prueba item decisorio",
                     Description = "Prueba item 2",
                     Duration = 10.0,
                     Order = 2,
                     AgendaItemType = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5",
                     RelatedDocumentsIds = [],
                     OrderType = "c2c12ed4-1001-466f-977c-07249f45c3a0",
                     ParentId = ""
                 };
                items.Add(item1);
                items.Add(item2);
                var results = await _eventAgendaItemService.AddOrUpdateAgendaItemsForEvent(ctx, BODY_ID, SHARED_EVENT_ID, items.ToArray());


                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task GetAgendaItems()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");

                var results = await _eventAgendaItemService.GetAgendaItemsByEvent(ctx, BODY_ID, SHARED_EVENT_ID);


                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


        [Fact]
        public async Task GetAgreements()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");

                var results = await _eventAgreementService.GetAgreementsByEvent(ctx, BODY_ID, SHARED_EVENT_ID);


                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task AddAgreements()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");

                var items = new EventAgreement()
                {
                    Id = "27",
                    Title = "Prueba acuerdo actualizado",
                    Description = "Prueba item 1",
                    Order = "3",
                    RelatedDocumentsIds = ["58a074ce-9ad8-4ed1-b5d8-20e81d26af28"],
                    RelatedCertificateId = null,
                    StatusId = "22abc0e5-25f1-4c07-80c5-b715dba4a1a6"
                };

                var results = await _eventAgreementService.AddOrUpdateAgreementsForEvent(ctx, BODY_ID, SHARED_EVENT_ID, [items]);


                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


        [Fact]
        public async Task AddRelatedDocumentToAgreement()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem($"/sites/{BODY_ID}");



                await _eventAgreementService.AddRemoveDocumentToAgreementByUniqueId(ctx, BODY_ID, SHARED_EVENT_ID, "26", "58a074ce-9ad8-4ed1-b5d8-20e81d26af28", true, true);
                var results = await _eventAgreementService.GetAgreementsByEvent(ctx, BODY_ID, SHARED_EVENT_ID);

                Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task AddEventsItems()
        {
            try
            {
                var jsonEvents = "[{\"Id\":27,\"Title\":\"Test desde back updated\",\"Description\":\"Test desde back\",\"Duration\":10,\"Order\":1,\"AgendaItemType\":\"c83ab3b1-00a1-4023-9a2f-1697b23c13d5\",\"RelatedDocumentsIds\":[]},{\"Id\":28,\"Title\":\"Test desde back 2\",\"Description\":\"Test desde back 2\",\"AgendaItemType\":\"c83ab3b1-00a1-4023-9a2f-1697b23c13d5\",\"Duration\":10,\"Order\":1,\"RelatedDocumentsIds\":[]}]";
                using var ctx = await CreatePnPContextAsUser("sites/aa-bb-ccc/");
                //var response = await eventsService.AddOrUpdateAgendaItemsForEventAsync(ctx, "23", JsonConvert.DeserializeObject<EventAgendaItem[]>(jsonEvents));

                //Assert.NotNull(response);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task DeleteEventItems()
        {
            try
            {
                using var ctx = await CreatePnPContextAsUser("sites/aa-bb-ccc/");

                // await eventsService.DeleteAgendaItem(ctx, "23", "28");

                // Assert.True(true);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }

        [Fact]
        public async Task SendRequestApprodvalsMinutesTest()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

                await _eventsService.SendRequestApprodvalsMinutes(ctx, BODY_ID, SHARED_EVENT_ID);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }





        [Fact]
        public async Task TestChangeDocuments()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem(new Uri($"/sites/aa-co-ccc").AbsolutePath);
                //var service = new EventCaptureService(this.Logger as ILogger<EventCaptureService>, ---falta PnPContextFactory---);

                //var diff = await service.CaptureAndprepareForSendingEventBySharedId(ctx, "aa-co-ccc", "14", "");
                //Assert.NotNull(diff);
                Assert.Fail();
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }



        [Fact]
        public async Task ChangeStatusEventTest()
        {
            try
            {
                using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);


                var diff = await _eventsService.ChangeStatusByEvent(ctx, BODY_ID, SHARED_EVENT_ID, DALConstants.TaxonomyValuesIds.EventStatus.Published);

                Assert.NotNull(diff);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


        [Fact]
        public async Task pendingChangeTest()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

                var test = await _eventsService.HasPendingChanges(ctx, BODY_ID, SHARED_EVENT_ID);

                Assert.NotNull(test);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


        [Fact]
        public async Task GetPublishDocuments()
        {
            try
            {
                using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
                var documentDalprovider = new EventPublishedDocumentDALprovider();

                var test = await documentDalprovider.GetPublishedDocument(ctx, SHARED_EVENT_ID);
                Assert.NotNull(test);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.ToString());
            }
        }


    }
}
