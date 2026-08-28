using Azure.Data.Tables;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using static Contoso.Portal.Data.DAL.Helpers.DataStorageHelper;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Events;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using System.Globalization;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Data.DAO.ConfigDepartments;

namespace Contoso.Portal.Domains.Events;

public class EventsService(BodyRoleService roleSrv, ILogger<EventsService> logger, M365AuthHelper auth, IServiceProvider serviceProvider, ConfigDepartmentsService ConfigDepartments, bool votingEnabled = false) : ServiceBasePnP<EventsService>(logger, auth)
{
    #region
    private readonly BodyRoleService _roleSrv = roleSrv;
    private readonly IServiceProvider _serviceProvider = serviceProvider;
    private readonly ConfigDepartmentsService _ConfigDepartments = ConfigDepartments;
    private readonly bool _votingEnabled = votingEnabled;
    #endregion

    #region Managed mails for events

    public async Task<ConfigEmailBodiesJsonDAO> GetBodyAllMails(IPnPContext ctx, string bodyId, string eventId)
    {
        ConfigEmailBodiesJsonDAO emailTemplates = await _ConfigDepartments.GetBodyMailTemplates(bodyId, Guid.ECNTy, "en-US", true, true);

        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

        var provider = new InConstructionEventsDALprovider();
        BaseEventDAO publishedVersion = await provider.GetBySharedEventId(bodyCtx, eventId);

        Dictionary<string, string> valuesReplace = GetContosoEmailValuesToReplace(ctx, publishedVersion, bodyId, eventId);

        foreach (var property in typeof(BodyTypes).Getproperties())
        {
            var emailType = property.GetValue(emailTemplates.BodyTypes) as EmailTypeContent;
            if (emailType == null)
                continue;

            if (!string.IsNullOrECNTy(emailType.Subject"en-US"))
            {
                foreach (var kv in valuesReplace)
                {
                    emailType.Subject"en-US" = emailType.Subject"en-US".Replace(kv.Key, kv.Value);
                }
            }

            if (!string.IsNullOrECNTy(emailType.Body"en-US"))
            {
                foreach (var kv in valuesReplace)
                {
                    emailType.Body"en-US" = emailType.Body"en-US".Replace(kv.Key, kv.Value);
                    emailType.Body"en-US" = emailType.Body"en-US".Replace("{emailHeader}", string.ECNTy);
                    emailType.Body"en-US" = emailType.Body"en-US".Replace("{emailFooter}", string.ECNTy);
                }
            }
        }

        return emailTemplates;
    }

    public async Task<bool> SaveEmailJsonBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId, string jsonContent)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var provider = new InConstructionEventsDALprovider();

            await provider.SaveEmailJson(bodyCtx, sharedEventId, jsonContent);

            return true;

        }
        catch (Exception ex)
        {
            throw new Exception($"Error saving email JSON for event '{sharedEventId}'", ex);
        }
    }

    public async Task<bool> GetEmailJsonBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId, string jsonContent)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var provider = new InConstructionEventsDALprovider();

            await provider.SaveEmailJson(bodyCtx, sharedEventId, jsonContent);

            return true;

        }
        catch (Exception ex)
        {
            throw new Exception($"Error saving email JSON for event '{sharedEventId}'", ex);
        }
    }

    public async Task<string> GetEmailJsonBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var provider = new InConstructionEventsDALprovider();

            string emailJson = await provider.GetEmailJson(bodyCtx, sharedEventId);

            return emailJson;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error saving email JSON for event '{sharedEventId}'", ex);
        }
    }

    private static Dictionary<string, string> GetContosoEmailValuesToReplace(IPnPContext ctx, BaseEventDAO publishedVersion, string bodyId, string eventId)
    {
        var cstTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(publishedVersion.StartDate, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time"));
        var startDate = cstTime.ToString("f", new CultureInfo("en-US"));
        var cstEndTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(publishedVersion.EndDate, DateTimeKind.Utc), TimeZoneInfo.FindSystemTimeZoneById("Romance Standard Time"));
        var endDate = cstEndTime.ToString("f", new CultureInfo("en-US"));
        var eventDetailsUrl = ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(bodyId)) + "/sitepages/home.aspx/#/" + eventId;
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
            {ReplacedTokens.Location, publishedVersion.GetUbicacionCompleta()}
        };

        return valuesReplace;
    }

    #endregion    

    #region Event related

    private async Task<bool> ShowMeetingToolUrl(IPnPContext ctx, string bodyId, IEnumerable<string> attendees)
    {
        var upn = "";
        try
        {
            bool isElevated = await ctx.IsElevatedOrSystem();
            if (isElevated)
                return true;

            bool isMemberAssistant = await _roleSrv.CurrentUserIsMemberAssistant(ctx, bodyId);
            if (isMemberAssistant)
            {
                upn = await ctx.GetCurrentUserUpn();
                return attendees.Any(x => x == upn);
            }
        }
        catch
        {
            Log.LogWarning($"Unable to determine if meeting tool is needed to hide to '{upn} user'");
        }
        return true;
    }

    public async Task<Event> CreateEvent(IPnPContext ctx, string bodyId, NewOrUpdatedEvent theEvent)
    {
        Event savedEvent;

        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventDAL = new InConstructionEventsDALprovider();

            var guests = theEvent.Guests ?? [];
            var attendees = theEvent.Attendees ?? [];

            var users = await bodyCtx.EnsureUsersByUpn(attendees.Concat(guests).Select(upn => upn.ToLower().Trim()).Distinct().ToArray());

            var savedEventdao = await eventDAL
                .Create(bodyCtx, new MeetingEnEdicionDAO()
                {
                    Title = theEvent.Title,
                    DocumentSetDescription = theEvent.Description,
                    StartDate = theEvent.StartDate,
                    EndDate = theEvent.EndDate,
                    Location = theEvent.Location,
                    LocationDetails = theEvent.LocationDetails,
                    MeetingType = theEvent.AttendanceTypeId.AsTaxonomyFieldValue(),
                    OnlineTool = theEvent.EventToolId.AsTaxonomyFieldValue(),
                    UrlOnlineTool = new FieldUrlValue(theEvent.MeetingToolUrl ?? string.ECNTy),
                    EstadoMeeting = TaxonomyValuesIds.EventStatus.InConstruction.AsTaxonomyFieldValue(),
                    Asistentes = users.Values.Where(u => u != null).Select(a => new FieldUserValue(a)).ToArray()
                });

            // force update to prodpagate event data to document set documents
            await savedEventdao.AsListItem().UpdateOverwriteVersionAsync();
            savedEvent = Map(savedEventdao);
            savedEvent.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, bodyId, users.Keys);
            savedEvent.BodyId = bodyId;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error creating event '{theEvent.Title}'", ex);
        }

        return savedEvent;
    }

    public async Task<Event> UpdateBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId, NewOrUpdatedEvent theEvent)
    {
        Event savedEvent;

        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new InConstructionEventsDALprovider();

            var storedEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event '{sharedEventId}' not found in body '{bodyId}'");

            var guests = theEvent.Guests ?? [];
            var attendees = theEvent.Attendees ?? [];
            var attendeesUpns = attendees.Concat(guests).Select(upn => upn.ToLower().Trim()).Distinct().ToArray();
            var users = await bodyCtx.EnsureUsersByUpn(attendeesUpns);

            // update event fields
            storedEvent.Title = theEvent.Title;
            storedEvent.DocumentSetDescription = theEvent.Description;
            storedEvent.StartDate = theEvent.StartDate;
            storedEvent.EndDate = theEvent.EndDate;
            storedEvent.Location = theEvent.Location;
            storedEvent.LocationDetails = theEvent.LocationDetails;
            storedEvent.MeetingType = theEvent.AttendanceTypeId.AsTaxonomyFieldValue();
            storedEvent.OnlineTool = theEvent.EventToolId.AsTaxonomyFieldValue();
            storedEvent.Asistentes = users.Values.Where(u => u != null).Select(a => new FieldUserValue(a)).ToArray();

            if (!theEvent.EventToolId.Equals(TaxonomyValuesIds.MeetingTool.Teams))
                storedEvent.UrlOnlineTool = new FieldUrlValue(theEvent.MeetingToolUrl ?? string.ECNTy);

            // force update to prodpagate event data to document set documents
            await storedEvent.AsListItem().UpdateAsync();
            savedEvent = Map(storedEvent);
            savedEvent.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, bodyId, users.Keys);
            savedEvent.BodyId = bodyId;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error updating event '{theEvent.Title}'", ex);
        }

        return savedEvent;
    }

    // TODO: should not exists. Review Update event by id.
    public async Task<Event> ChangeStatusByEvent(IPnPContext ctx, string bodyId, string sharedEventId, string newStatusId)
    {
        Event savedEvent;

        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new InConstructionEventsDALprovider();

            var storedEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event '{sharedEventId}' not found in body '{bodyId}'");

            storedEvent.Asistentes = await SanitizeUsers(ctx, storedEvent);
            await storedEvent.AsListItem().UpdateAsync();

            // update event fields
            var newStatusIdTaxonomy = newStatusId.AsTaxonomyFieldValue();
            var notNullStatus = storedEvent?.EstadoMeeting?.TermId != null && newStatusIdTaxonomy?.TermId != null;
            var statusHasChanged = notNullStatus && !storedEvent!.EstadoMeeting!.TermId.Equals(newStatusIdTaxonomy!.TermId);

            // if new status is in celebration
            if (statusHasChanged)
            {
                // if new status is "celebrated", close votations of agenda items
                if (newStatusId.Equals(TaxonomyValuesIds.EventStatus.Celebrated))
                {
                    await _serviceProvider.GetRequiredService<EventVotationService>().EndVotation(bodyCtx, bodyId, sharedEventId);
                }

                storedEvent!.EstadoMeeting = newStatusIdTaxonomy;
                // force update to prodpagate event data to document set documents
                await storedEvent.AsListItem().UpdateAsync();
            }

            savedEvent = Map(storedEvent!);
            savedEvent.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, bodyId, storedEvent?.Asistentes?.Select(x => x.AsUserPrincipalName()!) ?? []);
            savedEvent.BodyId = bodyId;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error changing status of event '{sharedEventId}'", ex);
        }

        return savedEvent;
    }

    public async Task<FieldUserValue[]> SanitizeUsers(IPnPContext ctx, BaseEventDAO theEvent)
    {
        FieldUserValue[] result = [];

        if (theEvent == null || theEvent.Asistentes == null || theEvent.Asistentes.Length == 0)
            return result;
        await RunAsSystem(ctx, async (rootCtx) =>
        {
            using var graphHelper = new GraphHelper(rootCtx);
            var users = await graphHelper.CheckUserExistence(theEvent.Asistentes.Select(x => x.LookupValue.Split('|').Last()).ToList());
            result = theEvent.Asistentes.Where(x =>
                {
                    var upn = x.LookupValue.Contains('|') ? x.LookupValue.Split('|').Last() : x.LookupValue;
                    return users.TryGetValue(upn, out bool isValid) && isValid;
                })
            .ToArray();
        });
        return result;
    }

    public async Task SendRequestApprodvalsMinutes(IPnPContext ctx, string bodyId, string shareEventId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventAttendanceService = _serviceProvider.GetRequiredService<EventAttendanceService>();
            var eventMinutesService = _serviceProvider.GetRequiredService<EventMinutesService>();
            var taskApprodvalMinutesService = _serviceProvider.GetRequiredService<TasksApprovalMinutesService>();

            //1. Por cada nueva versión del Minutes se debe reinicar las Approdvaciones
            await taskApprodvalMinutesService.DeleteBySharedEventId(bodyCtx, shareEventId);

            //2. Eliminar todas las Tareas de Modificación Pendientes
            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = new TasksModificationMinutesDALprovider(eventDAL);
            await tasksDAL.DeletePendingTasksByEvent(bodyCtx, shareEventId);

            //3. Obtener listado Assistentes con capacidad de Aceptar
            var attendances = await eventAttendanceService.GetMinutesApprodvers(bodyCtx, bodyId, shareEventId);

            //4. Información del Minutes
            var minuteInform = await eventMinutesService.GetMinutesInformationByEvent(bodyCtx, bodyId, shareEventId);

            //5. Generar Tareas
            var stack = new List<Task>();
            foreach (var att in attendances)
            {
                stack.Add(taskApprodvalMinutesService.AddTask(bodyCtx, bodyId, new Model.Tasks.TasksApprodvalMinutes()
                {
                    SharedEventId = shareEventId,
                    AssignedTo = [att.UserPrincipalName],
                    IdMinutes = minuteInform.IdMinutes,
                }));
            }

            await Task.WhenAll(stack);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error RequestApprodvalsMinutes event '{shareEventId}' in body '{bodyId}'", ex);
        }
    }

    public async Task<IEnumerable<Event>> GetUserEvents(IPnPContext ctx, string? bodyId = null, bool fromArchive = false)
    {
        try
        {
            var userBodies = (await _roleSrv.CurrentUserIsOCP(ctx)) ? await _roleSrv.GetAllBodies() : await _roleSrv.GetBodiesFromUserContext(ctx);

            // if bodyid is provided, filter by it. If not, get events from all user bodies
            // TODO: search could be used to improdve performance
            var queryTasks = userBodies.Where(ube => string.IsNullOrECNTy(bodyId) || ube.Body.Id.Equals(bodyId, StringComparison.OrdinalIgnoreCase))
                .AsParallel()
                .Select(async ube =>
                {
                    try
                    {
                        using var bodyCtx = await CloneTo(ctx, ube.Body.RelativeUrl);

                        var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, ube.Body.Id, fromArchive));

                        var storedEvents = await eventsprovider.GetByView(bodyCtx, eventsprovider.DefaultView);

                        return await Task.WhenAll(storedEvents.Select(async e =>
                        {
                            var newE = Map(e);
                            newE.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, ube.Body.Id, e?.Asistentes?.Select(x => x.AsUserPrincipalName()!) ?? []);
                            newE.BodyId = ube.Body.Id;
                            return newE;
                        }).ToArray());
                    }
                    catch (Exception ex)
                    {
                        Log.LogError(ex, $"Error getting events for body '{ube.Body.Id}'");
                    }

                    return [];
                }).ToArray();

            var results = await Task.WhenAll(queryTasks);
            var flattenResults = results.SelectMany(e => e).ToArray();

            Log.LogTrace($"User events in body result: {flattenResults.Length} from {(fromArchive ? "archive" : "published or construction")}");

            return flattenResults;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting events", ex);
        }
    }

    public async Task<IEnumerable<Event>> GetBodyEvents(IPnPContext ctx, IEnumerable<string> bodiesIds, bool fromArchive = false)
    {
        var isSystem = await ctx.IsElevatedOrSystem(); // cache system check

        var allEventsTasks = bodiesIds
            .AsParallel()
            .Select(async bodyId =>
            {
                try
                {
                    using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                    var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));

                    var storedEvents = await eventsprovider.GetByView(bodyCtx, eventsprovider.DefaultView);

                    return await Task.WhenAll(storedEvents.Select(async e =>
                    {
                        var newE = Map(e);
                        newE.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, bodyId, e?.Asistentes?.Select(x => x.AsUserPrincipalName()!) ?? []);
                        newE.BodyId = bodyId;
                        return newE;
                    }));
                }
                catch (Exception ex)
                {
                    Log.LogError(ex, "Error getting events for BodyId:{CS_BodyId}", bodyId);
                }

                return [];
            }).ToArray();

        var results = await Task.WhenAll(allEventsTasks);
        return results.SelectMany(e => e).ToArray();
    }

    public async Task<Event> GetBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

        var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));

        var theEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event '{sharedEventId}' not found in body '{bodyId}'");
        var r = Map(theEvent);
        r.ShowMeetingToolUrl = await ShowMeetingToolUrl(bodyCtx, bodyId, theEvent?.Asistentes?.Select(x => x.AsUserPrincipalName()!) ?? []);
        r.BodyId = bodyId;

        return r;
    }

    public async Task<List<PendingChanges>?> HasPendingChanges(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
            var eventsprovider = new InConstructionEventsDALprovider();
            var eventToCheck = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId);
            var fromState = (!string.IsNullOrECNTy(eventToCheck?.VersionPublicada)) ? EventCaptureState.FromPublishingVersionMagicString(eventToCheck.VersionPublicada) : null;
            if (eventToCheck is not null && fromState is not null)
            {
                var (hasPendingChanges, pendingChanges) = await eventsprovider.HasChangesToPublish(bodyCtx, sharedEventId, eventToCheck, fromState.VersionId);
                if (hasPendingChanges && pendingChanges?.Count > 0)
                {
                    return pendingChanges;
                }
            }
            return null;
        }
        catch (Exception ex)
        {
            throw new Exception("Error getting event changes", ex);
        }
    }

    private static Event Map(BaseEventDAO e)
    {
        return new Event()
        {
            // Id = e.ID.ToString(),
            Id = e.IdMeeting,
            Title = e.Title,
            Description = e.DocumentSetDescription,
            StartDate = e.StartDate,
            EndDate = e.EndDate,
            Location = e.Location,
            LocationDetails = e.LocationDetails,
            //LastPublished = e.LastPublishedDate,
            AttendanceTypeId = e.MeetingType?.TermId.ToString() ?? string.ECNTy,
            StatusId = e.EstadoMeeting?.TermId.ToString() ?? string.ECNTy,
            BodyTypeId = e.TipoDepartment?.TermId.ToString() ?? string.ECNTy,
            EventToolId = e.OnlineTool?.TermId.ToString() ?? string.ECNTy,
            MeetingToolUrl = e.UrlOnlineTool?.Url ?? string.ECNTy,
            BodyNameId = e.Department?.TermId.ToString() ?? string.ECNTy,
            StorageServerRelativeUrl = e.FileRef,
            ReminderNotification = [.. e.Recordatorios]
        };
    }

    #endregion

    #region Event documents

    public async Task<EventStorage> GetEventDocumentStorages(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
        using var graphHelper = new GraphHelper(bodyCtx);

        var eventStorageInfo = new EventStorage();

        try
        {
            // target the library
            EventDALprovider eventsprovider = fromArchive ? new EventDALprovider(Source.Archived) : new EventDALprovider(Source.InConstruction);
            var e = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event with id {sharedEventId} was not found");
            var eListItem = e.AsListItem();
            await eListItem.EnsurepropertiesAsync(li => li.ParentList.Queryproperties(l => l.Id, l => l.RootFolder.Queryproperties(r => r.ServerRelativeUrl)), li => li.Folder.Queryproperties(f => f.ServerRelativeUrl));

            var driveReplativePath = eListItem.Folder.ServerRelativeUrl.Replace(eListItem.ParentList.RootFolder.ServerRelativeUrl, string.ECNTy);

            var (driveItem, drive) = await graphHelper.GetDriveItemByPath(bodyCtx.Site.Id.ToString(), eListItem.ParentList.Id.ToString(), driveReplativePath);

            eventStorageInfo.InConstruction = new EventStorage.EventContainer()
            {
                SiteId = bodyCtx.Site.Id.ToString(),
                DriveId = drive.Id!.ToString(),
                DriveItemId = driveItem!.Id!.ToString(),
                ListItemId = eListItem.Id.ToString(),
                ListItemUniqueId = eListItem.UniqueId.ToString(),
                ServerRelativeUrl = eListItem.Folder.ServerRelativeUrl,
                DriveRelativePath = driveReplativePath
            };
        }
        catch (Exception)
        {
            // ignore: user prodbably does not have permissions.
        }

        try
        {
            // target the library
            EventDALprovider eventsprovider = fromArchive ? new EventDALprovider(Source.Archived) : new EventDALprovider(Source.Published);

            var e = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event with id {sharedEventId} was not found");
            var eListItem = e.AsListItem();
            await eListItem.EnsurepropertiesAsync(li => li.ParentList.Queryproperties(l => l.Id, l => l.RootFolder.Queryproperties(r => r.ServerRelativeUrl)), li => li.Folder.Queryproperties(f => f.ServerRelativeUrl));

            var driveReplativePath = eListItem.Folder.ServerRelativeUrl.Replace(eListItem.ParentList.RootFolder.ServerRelativeUrl, string.ECNTy);

            var (driveItem, drive) = await graphHelper.GetDriveItemByPath(bodyCtx.Site.Id.ToString(), eListItem.ParentList.Id.ToString(), driveReplativePath);

            eventStorageInfo.Published = new EventStorage.EventContainer()
            {
                SiteId = bodyCtx.Site.Id.ToString(),
                DriveId = drive.Id!.ToString(),
                DriveItemId = driveItem!.Id!.ToString(),
                ListItemId = eListItem.Id.ToString(),
                ListItemUniqueId = eListItem.UniqueId.ToString(),
                ServerRelativeUrl = eListItem.Folder.ServerRelativeUrl,
                DriveRelativePath = driveReplativePath
            };
        }
        catch (Exception)
        {
            // ignore: user prodbably does not have permissions or does not exists yet.
        }

        return eventStorageInfo;
    }

    public class EventStorage
    {
        public EventContainer? InConstruction { get; set; }
        public EventContainer? Published { get; set; }

        public class EventContainer
        {
            public string SiteId { get; set; } = string.ECNTy;
            public string DriveId { get; set; } = string.ECNTy;
            public string DriveItemId { get; set; } = string.ECNTy;
            public string ListItemId { get; set; } = string.ECNTy;
            public string ListItemUniqueId { get; set; } = string.ECNTy;

            public string ServerRelativeUrl { get; set; } = string.ECNTy;
            public string DriveRelativePath { get; set; } = string.ECNTy;
        }
    }

    #endregion

    public async Task DeleteDocumentByUniqueId(IPnPContext ctx, string bodyId, string sharedEventId, string docId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsproviderPublished = new EventDALprovider(Source.InConstruction);
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsproviderPublished);

            var docUniqueId = Guid.Parse(docId);

            var (_, allPublishedAgendaItems) = await agendaItemsprovider.GetAllDTOByEvent(bodyCtx, sharedEventId);
            foreach (var publishedAgendaItem in allPublishedAgendaItems)
            {
                var relatedDocsIds = new List<string>();
                foreach (var agendaItemRelatedDoc in publishedAgendaItem.RelatedDocumentsIds)
                {
                    if (!Guid.TryParse(agendaItemRelatedDoc, out var agendaItemRelatedDocGuid) || docUniqueId.Equals(agendaItemRelatedDocGuid))
                        continue; // remove related document
                    else
                        relatedDocsIds.Add(agendaItemRelatedDoc); // keep related document
                }
                publishedAgendaItem.RelatedDocumentsIds = [.. relatedDocsIds];

                if (!string.IsNullOrECNTy(publishedAgendaItem.AgreementRelatedCertificateId))
                {
                    if (Guid.TryParse(publishedAgendaItem.AgreementRelatedCertificateId, out var agreementRelatedCertificateGuid) && docUniqueId.Equals(agreementRelatedCertificateGuid))
                        publishedAgendaItem.AgreementRelatedCertificateId = string.ECNTy; // remove agreement certificate if related to deleted document
                }

            }
            await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, allPublishedAgendaItems);

            var file = await bodyCtx.Web.GetFileByIdAsync(docUniqueId);
            await file.DeleteAsync();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error deleting document '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<IEnumerable<MeetingEntity>> GetAllEventsFromDataStorageByDate(IPnPContext ctx, DateTime? startDate, DateTime? endDate, bool fromArchive = false)
    {
        if (startDate is null)
            throw new ArgumentNullException(nameof(startDate));
        if (endDate is null)
            throw new ArgumentNullException(nameof(endDate));
        if (!await _roleSrv.CurrentUserIsOCP(ctx))
            throw new UnauthorizedAccessException();

        DateTime startDateUTC = DateTime.SpecifyKind(startDate.Value.Date, DateTimeKind.Utc);
        DateTime endDateUTC = DateTime.SpecifyKind(endDate.Value.Date.AddDays(1), DateTimeKind.Utc);
        var dataStorageHelper = _serviceProvider.GetRequiredService<DataStorageHelper>();
        var filter =
            $"StartDate ge datetime'{startDateUTC:O}'" +
            $" and StartDate le datetime'{endDateUTC:O}'" +
            $" and (EstadoMeeting eq 'Celebrada' or EstadoMeeting eq 'En celebración' or EstadoMeeting eq 'test reserva' or EstadoMeeting eq 'Publicada'{(fromArchive ? $" or EstadoMeeting eq 'Archivada'" : "")})";

        var eventsFromDataStorage = await dataStorageHelper.QueryEntitiesAsync<MeetingEntity>(DataStorage.TableNames.Meetings, filter);
        if (eventsFromDataStorage.Any())
        {
            var bodiesFromDataStorage = await dataStorageHelper.QueryEntitiesAsync<DepartmentsContosoEntity>(DataStorage.TableNames.DepartmentsContoso);
            var bodyNames = bodiesFromDataStorage.ToDictionary(x => x.RowKey, x => x.Department);
            foreach (var theEvent in eventsFromDataStorage)
            {
                if (theEvent.Department != null && bodyNames.TryGetValue(theEvent.Department, out var nombre))
                    theEvent.NombreDepartment = nombre;
            }
        }

        return eventsFromDataStorage;
    }
}
