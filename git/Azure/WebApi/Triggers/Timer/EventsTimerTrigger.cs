using Microsoft.Azure.Functions.Worker;
using Microsoft.DurableTask;
using Microsoft.DurableTask.Client;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Runtime.Helpers;
using Contoso.Portal.Triggers.Http.EventsHttpTrigger;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Triggers.Timer;

public class EventsTimerTrigger(BodyRoleService bodyRoleService,
   NotificationsBusinessService notificationsBusinessService,
   EventCaptureService eventCaptureService,
   EventsService eventsService,
   EventExchangeSyncService eventExchangeService,
   ConfigDepartmentsService configBodyService,
   EventMinutesService eventMinutesService,
   TasksApprovalMinutesService tasksApprodvalMinutesService,
   ILogger<EventsTimerTrigger> logger,
   M365AuthHelper auth) : TriggerBase<EventsTimerTrigger>(logger, auth)
{
    private const int DEFAULT_DAYS_FOR_APprodVAL = 15;
    private readonly BodyRoleService _roleSrv = bodyRoleService;
    private readonly NotificationsBusinessService _notificationsBusinessService = notificationsBusinessService;
    private readonly EventCaptureService _eventCaptureService = eventCaptureService;
    private readonly EventsService _eventsService = eventsService;
    private readonly EventExchangeSyncService _eventExchangeService = eventExchangeService;
    private readonly EventMinutesService _eventMinutesService = eventMinutesService;
    private readonly TasksApprovalMinutesService _tasksApprodvalMinutesService = tasksApprodvalMinutesService;
    private readonly ConfigDepartmentsService _configBodyService = configBodyService;

    [Function(nameof(EventMonitoringJob))]
    public async Task EventMonitoringJob([TimerTrigger("0 */15 * * * *")] FunctionContext context, [DurableClient] DurableTaskClient client)
    {
        using var rootCtx = await CreatePnPContextAsSystem();
        var bodies = await _roleSrv.GetAllBodiesIds();

        var allPlatformEvents = (await _eventsService.GetBodyEvents(rootCtx, bodies))?.Where(e => e != null).ToArray() ?? [];

        foreach (var bodyEvents in allPlatformEvents.GroupBy(e => e.BodyId))
        {
            var bodyId = bodyEvents.Key;
            var events = bodyEvents.ToArray(); // all events

            try
            {
                // by status arrays
                var publishedEvents = events.Where((e) => e.StatusId.Equals(TaxonomyValuesIds.EventStatus.Published, StringComparison.OrdinalIgnoreCase)).ToArray();
                var inCelebrationEvents = events.Where((e) => e.StatusId.Equals(TaxonomyValuesIds.EventStatus.InCelebration, StringComparison.OrdinalIgnoreCase)).ToArray();
                var celebratedEvents = events.Where((e) => e.StatusId.Equals(TaxonomyValuesIds.EventStatus.Celebrated, StringComparison.OrdinalIgnoreCase)).ToArray();

                // events for notifications
                var eventsToNotifyStart24hoursByBody = publishedEvents
                    .Where(e => e.StartDate.ToUniversalTime() > DateTime.Now.ToUniversalTime() && e.StartDate.ToUniversalTime() < DateTime.Now.ToUniversalTime().AddDays(1)
                        && !e.ReminderNotification.Contains(NotificationsMessages.Reminders.TwentyFourHours))
                    .ToArray();
                var eventsToNotifyStartIn1hoursByBody = publishedEvents
                    .Where(e => e.StartDate.ToUniversalTime() > DateTime.Now.ToUniversalTime() && e.StartDate.ToUniversalTime() < DateTime.Now.ToUniversalTime().AddMinutes(60)
                        && !e.ReminderNotification.Contains(NotificationsMessages.Reminders.OneHour))
                    .ToArray();

                // events for status change
                var changeEventsToInCelebrationByBody = publishedEvents.Where(e => e.StartDate.ToUniversalTime() > DateTime.Now.ToUniversalTime().AddMinutes(-30) && e.StartDate.ToUniversalTime() < DateTime.Now.ToUniversalTime().AddMinutes(30)).ToArray();
                var changeEventsToCelebratedByBody = inCelebrationEvents.Where(e => e.EndDate.ToUniversalTime() < DateTime.Now.ToUniversalTime().AddHours(-24)).ToArray();

                if (changeEventsToInCelebrationByBody.Length != 0)
                {
                    Log.LogTrace($"Changing events to 'InCelebration' for BodyId:{{CS_BodyId}}: {string.Join("; ", changeEventsToInCelebrationByBody.Select(e => e.Id))}", bodyId);
                    foreach (var e in changeEventsToInCelebrationByBody)
                        await processEvent(nameof(changeEventsToInCelebrationByBody), bodyId, e, async (e) =>
                        {
                            await _eventsService.ChangeStatusByEvent(rootCtx, e.BodyId, e.Id, TaxonomyValuesIds.EventStatus.InCelebration);
                            var operation = await _eventCaptureService.CaptureAndprepareForSendingByEvent(rootCtx, e.BodyId, e.Id, "Automatic: change to 'In celebration'");
                            var instanceId = await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableEventSendingOrchestration), operation, new StartOrchestrationOptions() { InstanceId = $"es-b:{e.BodyId}-e:{e.Id}-c:{operation.ToCaptureState.CaptureId}" });
                        });
                }

                if (changeEventsToCelebratedByBody.Length != 0)
                {
                    Log.LogTrace($"Changing events to 'Celebrated' for BodyId:{{CS_BodyId}}: {string.Join("; ", changeEventsToCelebratedByBody.Select(e => e.Id))}", bodyId);
                    foreach (var e in changeEventsToCelebratedByBody)
                        await processEvent(nameof(changeEventsToCelebratedByBody), bodyId, e, async (e) =>
                        {
                            await _eventsService.ChangeStatusByEvent(rootCtx, e.BodyId, e.Id, TaxonomyValuesIds.EventStatus.Celebrated);
                            var operation = await _eventCaptureService.CaptureAndprepareForSendingByEvent(rootCtx, e.BodyId, e.Id, "Automatic: change to 'Celebrated'");
                            var instanceId = await client.ScheduleNewOrchestrationInstanceAsync(nameof(DurableEventSendingOrchestration), operation, new StartOrchestrationOptions() { InstanceId = $"es-b:{e.BodyId}-e:{e.Id}-c:{operation.ToCaptureState.CaptureId}" });
                        });
                }

                if (eventsToNotifyStart24hoursByBody.Length != 0)
                {
                    Log.LogTrace($"Sending 24h notifications for BodyId:{{CS_BodyId}}: {string.Join("; ", eventsToNotifyStart24hoursByBody.Select(e => e.Id))}", bodyId);
                    foreach (var e in eventsToNotifyStart24hoursByBody)
                        await processEvent(nameof(eventsToNotifyStart24hoursByBody), bodyId, e, async (e) =>
                         {
                             await _notificationsBusinessService.SendRememberEventNotification(rootCtx, e.BodyId, e.Id, RememberNotification.TwentyFourHours);
                         });
                }

                if (eventsToNotifyStartIn1hoursByBody.Length != 0)
                {
                    Log.LogTrace($"Sending 1h notifications for BodyId:{{CS_BodyId}}: {string.Join("; ", eventsToNotifyStartIn1hoursByBody.Select(e => e.Id))}", bodyId);
                    foreach (var e in eventsToNotifyStartIn1hoursByBody)
                        await processEvent(nameof(eventsToNotifyStartIn1hoursByBody), bodyId, e, async (e) =>
                        {
                            await _notificationsBusinessService.SendRememberEventNotification(rootCtx, e.BodyId, e.Id, RememberNotification.OneHour);
                        });
                }

                try
                {
                    // TODO: invert logic for better performance: get all minutes with startdateminutes instead of checking all celebrated events
                    int daysConfigApprodvalMinutes = await _configBodyService.GetPeriodDaysApprodvalMinutes(bodyId) ?? DEFAULT_DAYS_FOR_APprodVAL;

                    if (celebratedEvents.Length != 0)
                    {
                        Log.LogTrace($"Checking for pending minute approdval tasks for BodyId:{{CS_BodyId}}: {string.Join("; ", celebratedEvents.Select(e => e.Id))}", bodyId);
                        foreach (var e in celebratedEvents)
                        {
                            await processEvent("AutoAcceptPendingMinutesTasks", bodyId, e, async (e) =>
                            {
                                var statusMinutes = await _eventMinutesService.GetMinutesInformationByEvent(rootCtx, bodyId, e.Id);
                                if (statusMinutes.StartDateMinutes is null)
                                    return; // minutes approdval not started

                                var limitOfMinuteApprodvalDateInUtc = statusMinutes.StartDateMinutes.Value.AddDays(daysConfigApprodvalMinutes).ToUniversalTime();
                                var currentUtcTime = DateTime.Now.ToUniversalTime();

                                var limitHasPassed = DateTime.Compare(limitOfMinuteApprodvalDateInUtc, currentUtcTime) < 0;

                                if (limitHasPassed)
                                {
                                    await _tasksApprodvalMinutesService.AutoAcceptPendingMinutesTasks(rootCtx, bodyId, e.Id);
                                }
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.LogError(ex, "Error processing events --minutes-- for BodyId:{CS_BodyId}", bodyId);
                }
            }
            catch (Exception ex)
            {
                Log.LogError(ex, "Error processing events for BodyId:{CS_BodyId}", bodyId);
            }
        }

        async Task processEvent(string operation, string bodyId, Model.Events.Event theEvent, Func<Model.Events.Event, Task> action)
        {
            using var scope = Log.BeginContosoScope(operation, bodyId, theEvent.Id);
            try
            {
                await action(theEvent);
            }
            catch (Exception ex)
            {
                Log.LogError(ex, $"Error processing {operation} for BodyId:{{CS_BodyId}} EventId:{{CS_EventId}}", bodyId, theEvent.Id);
            }
        }
    }

    [Function(nameof(IdentifyExchangeAttendanceChanges))]
    public async Task IdentifyExchangeAttendanceChanges([TimerTrigger("0 */5 * * * *")] FunctionContext context, [DurableClient] DurableTaskClient client)
    {
        var result = await _eventExchangeService.processExchangeEventResponses();
    }
}