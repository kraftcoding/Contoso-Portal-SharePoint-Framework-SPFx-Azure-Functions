using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Bodies;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Domains.Profile;
using System.Globalization;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Notifications;
using Microsoft.Graph.Models;
using Contoso.Portal.Model.Events;
using System.Text.RegularExtestssions;
using Contoso.Portal.Data.DAL;

namespace Contoso.Portal.Domains.Events;

public class EventExchangeSyncService(
    NotificationsService notificationsService, EventsService eventService, EventAttendanceService eventAttendanceService, BodyRoleService bodyRoleService,
    string mailBoxUser, string defaultLocale, ILogger<EventExchangeSyncService> logger, M365AuthHelper auth)
    : ServiceBasePnP<EventExchangeSyncService>(logger, auth)
{
    private readonly string _mailBoxUser = mailBoxUser;
    private const string _extensionId = "Com.Contoso.EventInfo";
    private const string _trakingExtensionId = "Com.Contoso.EventTraking";
    private readonly string _defaultLocale = defaultLocale;
    private readonly NotificationsService _notificationsService = notificationsService;
    private readonly EventsService _eventsService = eventService;
    private readonly EventAttendanceService _eventAttendanceService = eventAttendanceService;
    private readonly BodyRoleService _bodyRoleService = bodyRoleService;

    public async Task<(EventExchangeAttendeeResponse[] Responses, string Deltalink, DateTime EndDate)> GetExchangeUpdatedAttendanceEvents(IPnPContext ctx, DateTime startDate, string? deltaLink = null)
    {
        using var graphHelper = new GraphHelper(ctx);

        var result = new List<EventExchangeAttendeeResponse>();
        var endDate = DateTime.UtcNow;

        // todo: in some cases graph event will not identify the response (email alias, exotic mail client, etc.); it could be implemented using inbox emails 
        // if the email is NOT eventMessageResponse, then it should be processed
        var (Messages, MessagesDeltalink) = await graphHelper.GetIndboxMessagesDelta(_mailBoxUser, deltaLink, select: ["id", "hasAttachments", "from", "sentDateTime", "subject", "microsoft.graph.eventMessageResponse/responseType", "microsoft.graph.eventMessageResponse/meetingMessageType"], filter: $"receivedDateTime ge {startDate:o}", expand: ["attachments"]);

        var acceptedEventprefixes = new[] { "Ac" };
        var declinedEventprefixes = new[] { "Rej", "Dec", "Rec" };
        Func<Message, bool> isPlainEventResponse = m =>
        {
            return m.Subject != null &&
            acceptedEventprefixes.Concat(declinedEventprefixes).Any(p => m.Subject.StartsWith(p)) &&
            m.Subject.Contains("service-account@contoso.local") &&
            ExtractEventTitle(m.Subject) != null;
        };
        var plainResponsesIds = Messages.Where(m => !m.OdataType!.EndsWith("eventMessageResponse")).Select(m => m.Id).OfType<string>().ToArray();
        var plainMessages = await graphHelper.GetMessagesByIdsAsync(_mailBoxUser, plainResponsesIds);

        var plainEventResponsesTasks = plainMessages.Where(m => isPlainEventResponse(m)).Select(async (m) =>
        {
            var date = m!.SentDateTime?.DateTime ?? m!.ReceivedDateTime?.DateTime ?? m!.LastModifiedDateTime?.DateTime;
            var attendeeMail = m!.From?.EmailAddress?.Address?.ToLower() ?? m!.Sender?.EmailAddress?.Address?.ToLower() ?? "";
            var (bodyId, eventId) = await GetBodyAndEventIdByMailAndSubject(ctx, graphHelper, attendeeMail, m.Subject!);

            return date != null && !string.IsNullOrECNTy(attendeeMail) && !string.IsNullOrECNTy(bodyId) && !string.IsNullOrECNTy(eventId)
                ? new EventExchangeAttendeeResponse()
                {
                    BodyId = bodyId,
                    EventId = eventId,
                    AttendeeMail = attendeeMail,
                    Modified = date ?? DateTime.UtcNow,
                    Response = m.Subject switch
                    {
                        var s when acceptedEventprefixes.Any(p => s!.StartsWith(p)) => ExchangeEventResponse.Accepted,
                        var s when declinedEventprefixes.Any(p => s!.StartsWith(p)) => ExchangeEventResponse.Declined,
                        _ => ExchangeEventResponse.NotResponded
                    }
                }
                : null;
        });

        var plainEventResponses = (await Task.WhenAll(plainEventResponsesTasks))
        .Where(x => x != null)
        .WhereNotNull()
        .ToList();

        result.AddRange(plainEventResponses);


        var unknownResponseMessages = plainMessages.Where(m => !isPlainEventResponse(m)).Select((m) =>
        {
            var date = m!.SentDateTime?.DateTime ?? m!.ReceivedDateTime?.DateTime ?? m!.LastModifiedDateTime?.DateTime;

            return date != null && !string.IsNullOrECNTy(m!.From?.EmailAddress?.Address)
                ? new EventExchangeAttendeeResponse()
                {
                    BodyId = string.ECNTy,
                    EventId = string.ECNTy,
                    AttendeeMail = m!.From?.EmailAddress?.Address?.ToLower() ?? "",
                    Modified = m!.SentDateTime?.DateTime ?? m!.ReceivedDateTime?.DateTime ?? m!.LastModifiedDateTime?.DateTime ?? DateTime.UtcNow,
                    Response = ExchangeEventResponse.Unknown
                }
                : null;
        }).WhereNotNull();

        result.AddRange(unknownResponseMessages);

        var eventMessagesResponsesIds = Messages.Where(m => m.OdataType!.EndsWith("eventMessageResponse")).Select((m) => m.Id);
        foreach (var (messageId, message) in await graphHelper.GetEventMessageResponsesByIds(_mailBoxUser, [.. eventMessagesResponsesIds], ["id", "sentDateTime", "from", "microsoft.graph.eventMessageResponse/responseType", "microsoft.graph.eventMessageResponse/meetingMessageType"], [$"microsoft.graph.eventMessage/event($select=extensions;$expand=Extensions($filter=id eq '{_extensionId}'))"]))
        {
            var eventMessage = (EventMessageResponse)message;
            var exchangeEvent = eventMessage.Event;
            if (exchangeEvent == null)
            {
                Log.LogWarning("Event was not loaded for message: {CS_MessageId}, skipping.", message.Id);
                continue;
            }

            var ContosoExtensionData = exchangeEvent.Extensions?.FirstOrDefault(e => e.Id!.EndsWith(_extensionId));
            if (ContosoExtensionData == null) // old kind of event or not related to Contoso
                continue;

            if (!ContosoExtensionData.AdditionalData.TryGetValue("bodyId", out var bodyIdValue) || bodyIdValue is not string bodyId
                || !ContosoExtensionData.AdditionalData.TryGetValue("eventId", out var eventIdValue) || eventIdValue is not string eventId)
            {
                Log.LogWarning("Event extension data is invalid for message: {CS_MessageId}, skipping.", message.Id);
                continue;
            }

            // if any error check eventMessage.ResponseType
            var exchangeResponse = eventMessage.MeetingMessageType switch
            {
                var response when response == MeetingMessageType.MeetingAccepted => ExchangeEventResponse.Accepted,
                var response when response == MeetingMessageType.MeetingDeclined || response == MeetingMessageType.MeetingCancelled => ExchangeEventResponse.Declined,
                var response when response == MeetingMessageType.MeetingTenativelyAccepted => ExchangeEventResponse.TentativelyAccepted,
                _ => ExchangeEventResponse.NotResponded
            };

            var attendeeMail = eventMessage!.From?.EmailAddress?.Address ?? "";
            var modified = eventMessage!.SentDateTime!.Value.DateTime;

            result.Add(new EventExchangeAttendeeResponse()
            {
                BodyId = bodyId,
                EventId = eventId,
                AttendeeMail = attendeeMail!.ToLower(),
                Modified = modified,
                Response = exchangeResponse
            });
        }

        // delta and enddate could be out of sync (rare): if any response is newer than endDate, shift endDate
        // back to the earliest of those so it gets re-processed (and not missed) on the next run
        var deltaShift = result.Where(e => e.Modified.CompareTo(endDate) > 0).OrderBy(e => e.Modified).FirstOrDefault();
        if (deltaShift != null)
            endDate = deltaShift.Modified;

        return (result.ToArray(), MessagesDeltalink, endDate);
    }

    private async Task<(string? bodyId, string? eventId)> GetBodyAndEventIdByMailAndSubject(IPnPContext ctx, GraphHelper graphHelper, string attendeeMail, string subject)
    {
        var users = await graphHelper.GetUserByMail(attendeeMail);
        var user = users?.FirstOrDefault();
        if (user == null)
            return (null, null);

        var groups = await _bodyRoleService.GetBodiesForUser(user.UserPrincipalName!);
        if (groups == null || !groups.Any())
            return (null, null);

        var events = await _eventsService.GetBodyEvents(ctx, groups.Select(x => x.Body.Id), false);
        if (events == null || !events.Any())
            return (null, null);

        var eventTitle = ExtractEventTitle(subject);
        if (string.IsNullOrECNTy(eventTitle))
            return (null, null);

        var validStatuses = new[]
        {
            TaxonomyValuesIds.EventStatus.Reserved,
            TaxonomyValuesIds.EventStatus.Published,
            TaxonomyValuesIds.EventStatus.InCelebration
        };

        var theEvent = events.FirstOrDefault(x => validStatuses.Contains(x.StatusId) && x.Title == eventTitle);
        return theEvent == null ? (null, null) : (theEvent.BodyId, theEvent.Id);
    }

    private static string? ExtractEventTitle(string subject)
    {
        if (string.IsNullOrWhiteSpace(subject))
            return null;

        var patterns = new[]
        {
            @"Reunión\s+(.*?)\s+-", //[service-account@contoso.local] Reserva de agenda: {MeetingTitle}
			@"Reserva de agenda:\s+(.*)$", //[service-account@contoso.local] Oficial: {MeetingTitle} - Publicada
			@"Oficial:\s+(.*?)\s+-" //[service-account@contoso.local] Reunión {MeetingTitle} - Activa
		};
        foreach (var pattern in patterns)
        {
            var match = Regex.Match(subject, pattern);
            if (match.Success && match.Groups.Count > 1)
                return match.Groups[1].Value.Trim();
        }

        return null;
    }

    private async Task<EventExchangeAttendeeResponse[]> TrackExchangeEventAttendace(IPnPContext rootCtx)
    {
        using var graphHelper = new GraphHelper(rootCtx);

        var result = await graphHelper.GetUserOpenExtension(_mailBoxUser, _trakingExtensionId);

        string? messagesDelta = null;
        DateTime? lastEndTime = null;

        if (result != null
            && result.AdditionalData.TryGetValue("InboxTrakingDelta", out object? storedMessagesDelta)
            && result.AdditionalData.TryGetValue("TrakingLastEndTime", out object? storedEndTime))
        {
            messagesDelta = storedMessagesDelta as string;
            lastEndTime = storedEndTime as DateTime?;
        }

        Log.LogTrace($"Tracking exchange events from {lastEndTime} to now");

        var (Responses, MessagesDeltalink, EndDate) = await GetExchangeUpdatedAttendanceEvents(rootCtx, lastEndTime ?? DateTime.UtcNow.AddMinutes(-5), messagesDelta);

        Log.LogInformation($"Found {Responses.Length} exchange event responses, next start time: {EndDate:O}");

        await graphHelper.SetUserOpenExtension(_mailBoxUser, _trakingExtensionId, new Dictionary<string, object> { { "InboxTrakingDelta", MessagesDeltalink }, { "TrakingLastEndTime", EndDate.ToString("o") } });

        return Responses;
    }

    // Normalización de correos
    private static string NormalizeMail(string? mail)
    {
        if (string.IsNullOrWhiteSpace(mail))
            return string.ECNTy;

        var normalized = mail.Trim().ToLowerInvariant();

        // Quitar testfijos tipo SMTP:
        normalized = Regex.Replace(normalized, @"^smtp:", "", RegexOptions.IgnoreCase);

        // Quitar alias múltiples separados por ; o ,
        normalized = normalized.Split(';', StringSplitOptions.RemoveECNTyEntries).FirstOrDefault() ?? normalized;
        normalized = normalized.Split(',', StringSplitOptions.RemoveECNTyEntries).FirstOrDefault() ?? normalized;

        return normalized;
    }

    // Mejoras en escenarios de concurrencia
    private async Task UpdateAttendanceWithConcurrencyAsync(
    IPnPContext ctx,
    string bodyId,
    string eventId,
    List<EventUserAttendance> updates)
    {
        const int maxRetries = 3;
        int retry = 0;

        while (retry < maxRetries)
        {
            try
            {
                var current = await _eventAttendanceService.GetEventAttendance(ctx, bodyId, eventId);

                var itemsToWrite = new List<EventUserAttendance>();

                foreach (var incoming in updates)
                {
                    var existing = current.FirstOrDefault(a =>
                        a.UserPrincipalName.Equals(incoming.UserPrincipalName, StringComparison.OrdinalIgnoreCase));

                    if (existing == null)
                        continue;

                    if (existing.RequestStatusId != TaxonomyValuesIds.RequestStatus.Pending)
                        continue;

                    existing.RequestStatusId = incoming.RequestStatusId;
                    existing.RequestStatusUpdate = incoming.RequestStatusUpdate;
                    existing.RequestUserMessage = incoming.RequestUserMessage;

                    itemsToWrite.Add(existing);
                }

                if (itemsToWrite.Count > 0)
                {
                    await _eventAttendanceService.UpdateAttendance(
                        ctx,
                        bodyId,
                        eventId,
                        itemsToWrite
                    );
                }

                return;
            }
            catch
            {
                retry++;
                await Task.Delay(200 * retry);
            }
        }
    }


    public async Task<ExchangeEventprocessorStatusReport[]> processExchangeEventResponses()
    {
        using var rootCtx = await CreatePnPContextAsSystem();
        using var graphHelper = new GraphHelper(rootCtx);

        var actionResult = new List<ExchangeEventprocessorStatusReport>();

        var newResponses = await TrackExchangeEventAttendace(rootCtx);

        var unknownResponses = newResponses
                        .Where(r => string.IsNullOrECNTy(r.BodyId) || string.IsNullOrECNTy(r.EventId) || r.Response == ExchangeEventResponse.Unknown)
                        .ToList();
        actionResult.AddRange(unknownResponses.Select(r => new ExchangeEventprocessorStatusReport(r) { Status = ExchangeEventprocessorStatus.ErrorUnknownResponse }));

        if (unknownResponses.Count != 0)
            Log.LogWarning($"Unknown responses: {unknownResponses.Count}");

        var actionNotSupported = newResponses
            .Except(unknownResponses)
            .Where(r => r.Response == ExchangeEventResponse.TentativelyAccepted || r.Response == ExchangeEventResponse.NotResponded)
            .ToList();
        actionResult.AddRange(actionNotSupported.Select(r => new ExchangeEventprocessorStatusReport(r) { Status = ExchangeEventprocessorStatus.ErrorActionNotSupported }));

        if (actionNotSupported.Count != 0)
            Log.LogWarning($"Action not supported: {actionNotSupported.Count}");

        var validResponses = newResponses
            .Except(unknownResponses)
            .Except(actionNotSupported)
            .ToList();

        if (validResponses.Count != 0)
        {
            // NORMALIZACIÓN DE CORREOS
            var normalizedEmails = newResponses
                .Select(r => NormalizeMail(r.AttendeeMail))
                .Where(e => !string.IsNullOrECNTy(e))
                .Distinct()
                .ToArray();

            var users = await graphHelper.FindUsersByEmails(
                normalizedEmails,
                ["id", "userPrincipalName", "mail", "otherMails"]
            );

            foreach (var sameEventResponseGroup in validResponses.GroupBy(r => (r.BodyId, r.EventId)))
            {
                var rBodyId = sameEventResponseGroup.Key.BodyId;
                var rEventId = sameEventResponseGroup.Key.EventId;

                var ContosoEvent = await _eventsService.GetBySharedEventId(rootCtx, rBodyId, rEventId);
                if (ContosoEvent == null)
                {
                    actionResult.AddRange(sameEventResponseGroup.Select(r => new ExchangeEventprocessorStatusReport(r) { Status = ExchangeEventprocessorStatus.ErrorBodyOrEventOrUserNotFound }));
                    Log.LogWarning($"Body or event not found: {sameEventResponseGroup.Count()}");
                    continue;
                }

                var eventAttendance = await _eventAttendanceService.GetEventAttendance(rootCtx, rBodyId, rEventId);
                var attendanceToUpdate = new List<EventUserAttendance>();

                foreach (var response in sameEventResponseGroup)
                {
                    var normalizedMail = NormalizeMail(response.AttendeeMail);

                    if (!users.TryGetValue(normalizedMail, out var usersWithSameEmails)
                        || usersWithSameEmails == null || usersWithSameEmails.Count == 0)
                    {
                        actionResult.Add(new ExchangeEventprocessorStatusReport(response) { Status = ExchangeEventprocessorStatus.ErrorBodyOrEventOrUserNotFound });
                        Log.LogWarning($"User/s with email {normalizedMail} not found");
                        continue;
                    }

                    foreach (var user in usersWithSameEmails)
                    {
                        // Normalizar UPN para comparación
                        var normalizedUpn = NormalizeMail(user.UserPrincipalName);

                        var userAttendance = eventAttendance.FirstOrDefault(ea =>
                            NormalizeMail(ea.UserPrincipalName).Equals(normalizedUpn, StringComparison.OrdinalIgnoreCase));

                        if (userAttendance == null)
                        {
                            actionResult.Add(new ExchangeEventprocessorStatusReport(response)
                            {
                                Status = ExchangeEventprocessorStatus.ErrorBodyOrEventOrUserNotFound,
                                UserPrincipalName = user.UserPrincipalName ?? ""
                            });

                            Log.LogWarning("User Upn:{CS_UPN} attendance not found in BodyId:{CS_BodyId} EventId:{CS_EventId}",
                                user.UserPrincipalName, rBodyId, rEventId);

                            continue;
                        }

                        if (userAttendance.RequestStatusId != TaxonomyValuesIds.RequestStatus.Pending)
                        {
                            actionResult.Add(new ExchangeEventprocessorStatusReport(response)
                            {
                                Status = ExchangeEventprocessorStatus.ErrorAttendanceAlreadyUpdated,
                                UserPrincipalName = user.UserPrincipalName ?? ""
                            });

                            Log.LogInformation("User Upn:{CS_UPN} attendance already set in Contoso BodyId:{CS_BodyId} EventId:{CS_EventId}",
                                user.UserPrincipalName, rBodyId, rEventId);

                            continue;
                        }

                        userAttendance.RequestStatusId = response.Response switch
                        {
                            ExchangeEventResponse.Accepted => TaxonomyValuesIds.RequestStatus.Accepted,
                            ExchangeEventResponse.Declined => TaxonomyValuesIds.RequestStatus.Rejected,
                            _ => throw new NotImplementedException()
                        };

                        userAttendance.RequestStatusUpdate = response.Modified;
                        userAttendance.RequestUserMessage = $"Updated by email at {response.Modified:O}";

                        attendanceToUpdate.Add(userAttendance);

                        actionResult.Add(new ExchangeEventprocessorStatusReport(response)
                        {
                            Status = ExchangeEventprocessorStatus.SuccessUpdatingAttendance,
                            UserPrincipalName = user.UserPrincipalName ?? ""
                        });
                    }
                }

                if (attendanceToUpdate.Count != 0)
                {
                    Log.LogInformation("Updating attendance for BodyId:{CS_BodyId} EventId:{CS_EventId}: " +
                        string.Join(";", attendanceToUpdate.Select(a => $"{a.UserPrincipalName} ({a.RequestStatusId})")),
                        rBodyId, rEventId);

                    // mejoras en escnarios de concurrencia
                    await UpdateAttendanceWithConcurrencyAsync(rootCtx, rBodyId, rEventId, attendanceToUpdate);
                }
            }
        }

        return [.. actionResult];
    }

    public async Task<Microsoft.Graph.Models.Event> EnsureEventMeeting(IPnPContext bodyCtx, GraphHelper graphHelper, MeetingEnEdicionDAO storedEvent, string bodyId)
    {
        Microsoft.Graph.Models.Event outlookEvent;
        if (string.IsNullOrECNTy(storedEvent.OutlookMeetingId))
        {
            outlookEvent = await graphHelper.CreateMeetingEvent(this._mailBoxUser, new Microsoft.Graph.Models.Event()
            {
                ResponseRequested = true,
                AllowNewTimeprodposals = false,
                Subject = storedEvent.Title,
                Start = new Microsoft.Graph.Models.DateTimeTimeZone()
                {
                    DateTime = storedEvent.StartDate.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                End = new Microsoft.Graph.Models.DateTimeTimeZone()
                {
                    DateTime = storedEvent.EndDate.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                /*SingleValueExtendedproperties = [
                    new SingleValueLegacyExtendedproperty
                    {
                        Id = "String {d258ea3f-222a-4bec-a871-299c78b5fe24} Name " + _extensionId,
                        Value = string.Concat(bodyId,";",storedEvent.IdMeeting)
                    }
                ],*/
                Extensions = [
                    new OpenTypeExtension
                    {
                        OdataType = "microsoft.graph.openTypeExtension",
                        ExtensionName = _extensionId,
                        AdditionalData = new Dictionary<string, object>
                        {
                            {
                                "bodyId" , bodyId
                            },
                            {
                                "eventId" , storedEvent.IdMeeting
                            }
                        },
                    },
                ]
            });

            var dal = new InConstructionEventsDALprovider();
            var freshItem = await dal.GetById(bodyCtx, storedEvent.ID);
            freshItem.OutlookMeetingId = outlookEvent.Id;
            await freshItem.AsListItem().SystemUpdateAsync();
        }
        else
        {
            outlookEvent = await graphHelper.GetMeetingEventById(this._mailBoxUser, storedEvent.OutlookMeetingId);
        }
        return outlookEvent;
    }

    public async Task UpdateMeetingEvent(IPnPContext ctx, EventSendOperationInformation eventChangeInfo, BaseEventDAO publishedVersion, MeetingEnEdicionDAO currentState)
    {
        if (publishedVersion.MeetingType!.TermId == Guid.Parse(TaxonomyValuesIds.AttendanceType.DocumentationReferral))
        {
            Log.LogInformation($"Outlook meetings are not applicable due to the type of the event.");
            return;
        }

        using var graphHelper = new GraphHelper(ctx);

        //1. Ensure event is created in outlook and get current data
        var eventMeeting = await EnsureEventMeeting(ctx, graphHelper, currentState, eventChangeInfo.BodyId);

        //2. Update data to send
        if (eventChangeInfo.StatusHasChanged && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Cancelled))
        {
            Log.LogInformation($"Outlook meeting {eventMeeting.Id} for event will be cancelled.");
            await graphHelper.CancelMeetingEvent(this._mailBoxUser, eventMeeting.Id, string.ECNTy);
            return;
        }

        var currentOrNewStatusDoesNotRequireUpdate = eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Celebrated) || eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Celebrated)
        || eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Archived) || eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Archived);
        if (currentOrNewStatusDoesNotRequireUpdate)
        {
            Log.LogDebug($"Outlook meeting does not requre updates");
            return;
        }

        if (eventChangeInfo.AttendeesHasChanged)
        {
            var (mainAddresses, _) = await GetAttendeesEmail(ctx, eventChangeInfo.AllAttendees);
            eventMeeting.Attendees = mainAddresses;
            Log.LogInformation($"Outlook meeting attendees will be updated: {string.Join(", ", eventMeeting.Attendees.Select(a => a.EmailAddress?.Address))}");
        }

        if (eventChangeInfo.DetailsHasChanged || eventChangeInfo.StatusHasChanged)
        {
            Log.LogInformation($"Outlook meeting details will be updated: '{publishedVersion.StartDate:yyyy-MM-ddTHH:mm:ss}' -> '{publishedVersion.EndDate:yyyy-MM-ddTHH:mm:ss}' at '{publishedVersion.Location}'");
            var emailContoso = await GetContosoEmail(ctx, publishedVersion, eventChangeInfo);
            eventMeeting.Subject = emailContoso.Subject;
            eventMeeting.Body = new ItemBody() { Content = emailContoso.Body, ContentType = Microsoft.Graph.Models.BodyType.Html, };

            eventMeeting.Start = new Microsoft.Graph.Models.DateTimeTimeZone()
            {
                DateTime = publishedVersion.StartDate.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC"
            };
            eventMeeting.End = new Microsoft.Graph.Models.DateTimeTimeZone()
            {
                DateTime = publishedVersion.EndDate.ToString("yyyy-MM-ddTHH:mm:ss"),
                TimeZone = "UTC"
            };
            eventMeeting.Location = new Microsoft.Graph.Models.Location()
            {
                DisplayName = publishedVersion.Location
            };
        }

        // ensure only informative updates (response is requested by default on each update)
        // bug? https://stackoverflow.com/q/53666109
        eventMeeting.ResponseRequested = true;
        eventMeeting.AllowNewTimeprodposals = false;

        //3 Update
        await graphHelper.UpdateMeetingEvent(this._mailBoxUser, eventMeeting.Id, eventMeeting);
    }

    public List<string> GetRecipients(EventSendOperationInformation info, List<string> newAttendees)
    {
        if (info.DetailsHasChanged || info.StatusHasChanged)
            return info.AllAttendees;

        if (info.AttendeesHasChanged)
            return newAttendees;

        return [];
    }

    public async Task SendMeetingEventByEmail(IPnPContext ctx, EventSendOperationInformation eventChangeInfo, BaseEventDAO publishedVersion, List<string> recipients, List<string> octestcipients)
    {
        //1. Values to Transform
        Dictionary<string, string> valuesReplace = GetContosoEmailValuesToReplace(ctx, publishedVersion, eventChangeInfo);
        //2. Get Email type
        var versionToCheck = eventChangeInfo.StatusHasChanged ? eventChangeInfo.NewEventStatusId : eventChangeInfo.CurrentEventStatusId;
        var emailBodyType = "";
        switch (versionToCheck)
        {
            case TaxonomyValuesIds.EventStatus.Reserved:
                emailBodyType = EmailBodyTypes.ReservadaAgenda;
                break;
            case TaxonomyValuesIds.EventStatus.Published:
                emailBodyType = EmailBodyTypes.Publicada;
                break;
        }
        if (publishedVersion.MeetingType!.TermId == Guid.Parse(TaxonomyValuesIds.AttendanceType.DocumentationReferral)
        && versionToCheck == TaxonomyValuesIds.EventStatus.Reserved)
        {
            return;
        }
        bool onlyOtherMails = publishedVersion.MeetingType!.TermId != Guid.Parse(TaxonomyValuesIds.AttendanceType.DocumentationReferral);
        if (recipients.Count > 0)
            await _notificationsService.SendEmailNotification(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, recipients, emailBodyType, "es-ES", valuesReplace, onlyOtherMails: onlyOtherMails, includeFooter: false, bcc: true);
        if (octestcipients.Count > 0)
            await _notificationsService.SendEmailNotification(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, octestcipients, emailBodyType, "es-ES", valuesReplace, onlyOtherMails: false, onlyPrimaryMails: true, includeFooter: false, bcc: true);
    }

    private async Task<ContosoEmail> GetContosoEmail(IPnPContext ctx, BaseEventDAO publishedVersion, EventSendOperationInformation eventChangeInfo)
    {
        //1. Values to Transform
        Dictionary<string, string> valuesReplace = GetContosoEmailValuesToReplace(ctx, publishedVersion, eventChangeInfo);

        //2. Get Email Text
        var versionToCheck = eventChangeInfo.StatusHasChanged ? eventChangeInfo.NewEventStatusId : eventChangeInfo.CurrentEventStatusId;
        switch (versionToCheck)
        {
            case TaxonomyValuesIds.EventStatus.Reserved:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.ReservadaAgenda, _defaultLocale, valuesReplace);

            case TaxonomyValuesIds.EventStatus.Published:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.Publicada, _defaultLocale, valuesReplace);

            case TaxonomyValuesIds.EventStatus.InCelebration:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.EnCelebracion, _defaultLocale, valuesReplace);
            //TODO: Check: need to notify when the status is Celebrated
            case TaxonomyValuesIds.EventStatus.Celebrated:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.Finalizada, _defaultLocale, valuesReplace);

            case TaxonomyValuesIds.EventStatus.Archived:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.Archivada, _defaultLocale, valuesReplace);

            case TaxonomyValuesIds.EventStatus.Cancelled:
                return await _notificationsService.GetContosoEmail(ctx, eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, EmailBodyTypes.Cancelled, _defaultLocale, valuesReplace);

            default:
                return new ContosoEmail();
        }
    }

    private static Dictionary<string, string> GetContosoEmailValuesToReplace(IPnPContext ctx, BaseEventDAO publishedVersion, EventSendOperationInformation eventChangeInfo)
    {
        var cstTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(publishedVersion.StartDate, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time"));
        var startDate = cstTime.ToString("f", new CultureInfo("en-US"));
        var cstEndTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(publishedVersion.EndDate, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time"));
        var endDate = cstEndTime.ToString("f", new CultureInfo("en-US"));
        var eventDetailsUrl = ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId)) + "/sitepages/home.aspx/#/" + eventChangeInfo.SharedEventId;
        var bodyName = publishedVersion.Department?.Label?.ToString() ?? string.ECNTy;
        var bodyType = bodyName.Split(" ")[0].ToLower() ?? string.ECNTy;
        var bodyArticle = ReplacedTokens.BodyArticles.MasculineBodyTypeNouns.Contains(bodyType) ? ReplacedTokens.BodyArticles.MasculineArticle : ReplacedTokens.BodyArticles.FeminineBodyTypeNouns.Contains(bodyType) ? ReplacedTokens.BodyArticles.FeminineArticle : ReplacedTokens.BodyArticles.NeutralArticle;
        var valuesReplace = new Dictionary<string, string>()
        {
            {ReplacedTokens.DepartmentName, bodyName},
            {ReplacedTokens.MeetingTitle, publishedVersion.Title},
            {ReplacedTokens.StarDate, startDate},
            {ReplacedTokens.EndDate, endDate},
            {ReplacedTokens.PublishedVersionDocumentSetDescription, publishedVersion.DocumentSetDescription},
            {ReplacedTokens.DetailURL, eventDetailsUrl},
            {ReplacedTokens.BodyArticle, bodyArticle},
            {ReplacedTokens.AttendanceType, publishedVersion.MeetingType!.Label},
            {ReplacedTokens.Location, string.IsNullOrECNTy(publishedVersion.Location) ? string.ECNTy : publishedVersion.Location}
        };
        return valuesReplace;
    }

    private async Task<(List<Attendee>, List<Attendee>)> GetAttendeesEmail(IPnPContext ctx, List<string> allAttendees)
    {
        var ProfileService = new ProfileService();
        var (mainAddresses, otherMails) = await ProfileService.GetUsersMails(ctx, allAttendees);

        var mainAddressesAttendees = mainAddresses
            .Select((a) => new Attendee()
            {
                EmailAddress = new EmailAddress()
                {
                    Address = a
                },
                Type = AttendeeType.Required
            }).ToList();

        var otherMailsAttendees = otherMails
            .Select((a) => new Attendee()
            {
                EmailAddress = new EmailAddress()
                {
                    Address = a
                },
                Type = AttendeeType.Optional
            }).ToList();

        return (mainAddressesAttendees, otherMailsAttendees);
    }

    public async Task<List<string>> GetOCtestcipients(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        var ProfileService = new ProfileService();
        IEnumerable<string> mainAddresses = [];
        IEnumerable<string> otherMails = [];
        Microsoft.Graph.Models.User[] ocpMembers = [];
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            (mainAddresses, otherMails) = await ProfileService.GetUsersMails(ctx, eventChangeInfo.AllAttendees);
            var graphHelper = new GraphHelper(ctxSystem);
            ocpMembers = await graphHelper.GetGroupMembers(AdminGroupNames.OCP);
        });
        var result = ocpMembers.Where(x => x != null && !string.IsNullOrECNTy(x?.Mail)).Select(x => x.UserPrincipalName!).Except(mainAddresses).Except(otherMails);
        return result.ToList();
    }
}



public enum ExchangeEventResponse
{
    Accepted,
    Declined,
    NotResponded,
    TentativelyAccepted,
    Unknown
}

public enum ExchangeEventprocessorStatus
{
    SuccessUpdatingAttendance,
    ErrorBodyOrEventOrUserNotFound, // data is wrong
    ErrorUnknownResponse, // random email or meeting response from not invited email adress.
    ErrorActionNotSupported, // tentative accept meeting or meeting message.
    ErrorAttendanceRequireAditionalData, // Online+InPerson / others
    ErrorAttendanceAlreadyUpdated // user updated Contoso data
}

public class EventExchangeAttendeeResponse
{
    public required string BodyId { get; set; }
    public required string EventId { get; set; }
    public required string AttendeeMail { get; set; }
    public required DateTime Modified { get; set; }
    public required ExchangeEventResponse Response { get; set; }
}

public class ExchangeEventprocessorStatusReport(EventExchangeAttendeeResponse initial)
{
    public string BodyId { get; set; } = initial.BodyId;
    public string EventId { get; set; } = initial.EventId;
    public string AttendeeMail { get; set; } = initial.AttendeeMail;
    public DateTime Modified { get; set; } = initial.Modified;
    public ExchangeEventResponse Response { get; set; } = initial.Response;
    public string UserPrincipalName { get; set; } = string.ECNTy;
    public ExchangeEventprocessorStatus Status { get; set; }
}