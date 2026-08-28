using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Domains.Events;
using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Notifications;

using Contoso.Portal.Common;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Runtime;
using Contoso.Portal.Model.Events;

using static Contoso.Portal.Data.DAL.DALConstants;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Extensions;
using System.Linq;
using System.Collections.Generic;
using Contoso.Portal.Data.DAL.Event;
using Azure.Data.Tables;
using Azure;
using Azure.Identity;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography.X509Certificates;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Data.DAL.Tasks;
using PnP.Core.Services;

namespace WebApi.Tests;

public class UtilityTestingFunctionTests : BasePnPAppTest
{



    public UtilityTestingFunctionTests(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output) { }

    [Fact]
    public async Task Test()
    {
        using var ctx = await CreatePnPContextAsSystem();
        var graphHelper = new GraphHelper(ctx);

        var settings = ServiceProvider().GetRequiredService<Settings>();
        var dataStorageHelper = ServiceProvider().GetRequiredService<DataStorageHelper>();
        var managementService = ServiceProvider().GetRequiredService<ManagementService>();
        var bodyRoleService = ServiceProvider().GetRequiredService<BodyRoleService>();
        var bodiesService = ServiceProvider().GetRequiredService<BodiesService>();
        var ConfigDepartmentsService = ServiceProvider().GetRequiredService<ConfigDepartmentsService>();
        var eventsService = ServiceProvider().GetRequiredService<EventsService>();
        var eventAgendaItemsService = ServiceProvider().GetRequiredService<EventAgendaItemsService>();
        var eventVotationService = ServiceProvider().GetRequiredService<EventVotationService>();
        var eventMinutesService = ServiceProvider().GetRequiredService<EventMinutesService>();
        var eventMinutesTemplateService = ServiceProvider().GetRequiredService<EventMinutesTemplateService>();
        var eventAttendanceService = ServiceProvider().GetRequiredService<EventAttendanceService>();
        var eventCaptureService = ServiceProvider().GetRequiredService<EventCaptureService>();
        var eventAgreementService = ServiceProvider().GetRequiredService<EventAgreementService>();
        var ProfileService = ServiceProvider().GetRequiredService<ProfileService>();
        var tasksService = ServiceProvider().GetRequiredService<TasksService>();
        var tasksAttendanceService = ServiceProvider().GetRequiredService<TasksAttendanceService>();
        var eventAttendanceTemplateService = ServiceProvider().GetRequiredService<EventAttendanceTemplateService>();
        var tasksDelegateService = ServiceProvider().GetRequiredService<TasksDelegateService>();
        var tasksApprodvalMinutesService = ServiceProvider().GetRequiredService<TasksApprovalMinutesService>();
        var tasksModificationMinutesService = ServiceProvider().GetRequiredService<TasksModificationMinutesService>();
        var tasksCertificationService = ServiceProvider().GetRequiredService<TasksCertificationService>();
        var notificationsBusinessService = ServiceProvider().GetRequiredService<NotificationsBusinessService>();
        var notificationsService = ServiceProvider().GetRequiredService<NotificationsService>();
        var m365AuthHelper = ServiceProvider().GetRequiredService<M365AuthHelper>();
        var taxonomyTranslationService = ServiceProvider().GetRequiredService<TaxonomyTranslationService>();
        var eventPublishingService = ServiceProvider().GetRequiredService<EventPublishingService>();
        var simService = ServiceProvider().GetRequiredService<INotificationSmsService>();
        var insideService = ServiceProvider().GetRequiredService<IDocumentArchiveService>();
        var outlookSyncService = ServiceProvider().GetRequiredService<EventExchangeSyncService>();
        var powerBiReportsService = ServiceProvider().GetRequiredService<PowerBiReportsService>();

        try
        {
            var bodyId = "demoev-co-cna";
            var token = "";
            var context = await CreatePnPContextAsUser($"sites/{bodyId}", token);
            //var events = await eventsService.GetUserEvents(context, bodyId, false);
            // var groups = await bodyRoleService.GetUserGroupsNames(userUPN);
            // var allBodies = await bodyRoleService.GetAllBodiesIds();
            var bodies = await ConfigDepartmentsService.GetConfigDepartments();
            var body = bodies.Where(t => t.Title.Equals(bodyId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            // CREATE EVENT
            // ONLINE EVENT
            var newOnlineEvent = new NewOrUpdatedEvent()
            {
                Title = $"Test {nameof(TaxonomyValuesIds.AttendanceType.Online)} 1",
                AttendanceTypeId = TaxonomyValuesIds.AttendanceType.Online,
                Description = $"Meeting {nameof(TaxonomyValuesIds.AttendanceType.Online)}",
                StartDate = DateTime.Now.AddHours(1),
                EndDate = DateTime.Now.AddHours(2),
                Location = "",
                LocationDetails = "",
                EventToolId = TaxonomyValuesIds.MeetingTool.Teams,
                Attendees = ["dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com", "contoso-testadmin@contoso-test.onmicrosoft.com"],
                Guests = []
            };

            var onlineEvent = await eventsService.CreateEvent(ctx, bodyId, newOnlineEvent);

            var agendaItemsResult = await eventAgendaItemsService.AddOrUpdateAgendaItemsForEvent(ctx, bodyId, onlineEvent.Id, [new()
                {
                    Title = "Acuerdo",
                    AgendaItemType = TaxonomyValuesIds.AgendaItemType.ForDecision,
                    Description = "Descripcion acuerdo",
                    Duration = 60
                }
            ]);
            await eventsService.ChangeStatusByEvent(ctx, bodyId, onlineEvent.Id, TaxonomyValuesIds.EventStatus.Published);
            await SendChanges(bodyId, onlineEvent.Id);

            // InPerson
            var newInPersonEvent = new NewOrUpdatedEvent()
            {
                Title = $"Test {nameof(TaxonomyValuesIds.AttendanceType.InPerson)} 1",
                AttendanceTypeId = TaxonomyValuesIds.AttendanceType.InPerson,
                Description = $"Meeting {nameof(TaxonomyValuesIds.AttendanceType.InPerson)}",
                StartDate = DateTime.Now.AddHours(1),
                EndDate = DateTime.Now.AddHours(2),
                Location = "Madrid",
                LocationDetails = "Calle de la Piruleta, 9",
                EventToolId = string.ECNTy,
                Attendees = ["dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com", "contoso-testadmin@contoso-test.onmicrosoft.com"],
                Guests = []
            };

            var InPersonEvent = await eventsService.CreateEvent(ctx, bodyId, newInPersonEvent);

            var agendaItemsResultInPerson = await eventAgendaItemsService.AddOrUpdateAgendaItemsForEvent(ctx, bodyId, InPersonEvent.Id, [new()
                {
                    Title = "Acuerdo",
                    AgendaItemType = TaxonomyValuesIds.AgendaItemType.ForDecision,
                    Description = "Descripcion acuerdo",
                    Duration = 60
                }
            ]);
            await eventsService.ChangeStatusByEvent(ctx, bodyId, InPersonEvent.Id, TaxonomyValuesIds.EventStatus.Published);
            await SendChanges(bodyId, InPersonEvent.Id);

            // REMISION DOCUMENTACION
            var newDocumentationSubmission = new NewOrUpdatedEvent()
            {
                Title = $"Test {nameof(TaxonomyValuesIds.AttendanceType.DocumentationReferral)} 1",
                AttendanceTypeId = TaxonomyValuesIds.AttendanceType.DocumentationReferral,
                Description = $"Meeting {nameof(TaxonomyValuesIds.AttendanceType.DocumentationReferral)}",
                StartDate = DateTime.Now.AddHours(1),
                EndDate = DateTime.Now.AddHours(2),
                Location = "",
                LocationDetails = "",
                EventToolId = string.ECNTy,
                Attendees = ["dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com", "contoso-testadmin@contoso-test.onmicrosoft.com"],
                Guests = []
            };

            var documentationSubmissionEvent = await eventsService.CreateEvent(ctx, bodyId, newDocumentationSubmission);

            var agendaItemsResultDocumentationSubmission = await eventAgendaItemsService.AddOrUpdateAgendaItemsForEvent(ctx, bodyId, documentationSubmissionEvent.Id, [new()
                {
                    Title = "Acuerdo",
                    AgendaItemType = TaxonomyValuesIds.AgendaItemType.ForDecision,
                    Description = "Descripcion acuerdo",
                    Duration = 60
                }
            ]);
            await eventsService.ChangeStatusByEvent(ctx, bodyId, documentationSubmissionEvent.Id, TaxonomyValuesIds.EventStatus.Published);
            await SendChanges(bodyId, documentationSubmissionEvent.Id);

            Debug.WriteLine("prodCESO FINALIZADO");


            // UPDATE EVENT THINGS: Replace for the already created event ID and comment the existant 'theEvent' variable
            // var theEvent = new
            // {
            //     Id = "fe0a91203b9a4954a2e51799a69697dc"
            // };

            //   EVENT prodPERTIES
            // await eventsService.UpdateBySharedEventId(ctx, bodyId, theEvent.Id, newOrUpdatedEvent);

            //   ATTENDANCE
            // await eventAttendanceService.AddOrRemoveEventAttendee(ctx, bodyId, theEvent.Id, "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com", true);
            // await eventAttendanceService.AddOrRemoveEventAttendee(ctx, bodyId, theEvent.Id, "fnietolo_emeal.nttdata.com@contoso-test.onmicrosoft.com", true);
            // await eventAttendanceService.AddOrRemoveEventAttendee(ctx, bodyId, theEvent.Id, "CNTextmiembro1@contoso-test.onmicrosoft.com", true);


            //   AGENDA ITEMS


            //   EVENT STATUS
            // await eventsService.ChangeStatusByEvent(ctx, bodyId, theEvent.Id, TaxonomyValuesIds.EventStatus.Reserved);

            //ALWAYS: SEND CHANGES TO prodPAGATE FOR ALL USERS

            async Task<EventSendOperationInformation?> SendChanges(string bodyId, string eventId)
            {
                EventSendOperationInformation? captureResult;
                try
                {
                    captureResult = await eventCaptureService.CaptureAndprepareForSendingByEvent(ctx, bodyId, eventId, "Hola hola !!");
                    // captureResult.SkipUserCommunications = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    throw;
                }

                try
                {
                    await eventPublishingService.EventCaptureToDocumentSetUpdates(ctx, captureResult);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    throw;
                }

                try
                {
                    await eventPublishingService.EventCaptureToTasksUpdates(ctx, captureResult);
                    await eventPublishingService.EventCaptureToTeamsAndOutlookUpdates(ctx, captureResult);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }

                try
                {
                    await eventPublishingService.EventCaptureToNotifications(ctx, captureResult);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }

                try
                {
                    await eventPublishingService.EventCaptureToRemovalOrArchive(ctx, captureResult);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                    throw;
                }

                return captureResult;
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine(ex.Message);
            throw;
        }
    }
}
