using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Extensions;
using PnP.Core.Model;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Data.DAO;

using Contoso.Portal.Model.Events;
using Contoso.Portal.Data.Data.DAL.Helpers;
using System.Text.RegularExtestssions;
using System.Globalization;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DTO.Event;

namespace Contoso.Portal.Domains.Events;

public partial class EventPublishingService : ServiceBasePnP<EventPublishingService>
{
    // services
    private readonly BodyRoleService _roleSrv;
    private readonly NotificationsService _notificationsService;
    private readonly TasksService _taskService;
    private readonly IDocumentArchiveService _insideService;
    private readonly ConfigDepartmentsService _configBodyService;
    private readonly EventExchangeSyncService _outlookSyncService;

    private readonly string _mailBoxUser;
    private readonly string _defaultLocale;
    private readonly Guid _tenantId;
    private readonly bool _sensivityReplicationEnabled;
    private readonly bool _pdfConversionEnabled;

    public EventPublishingService(
        BodyRoleService bodyRoleService, NotificationsService notificationsService, TasksService tasksService, IDocumentArchiveService insideService, ConfigDepartmentsService configBodyService, EventExchangeSyncService outlookSyncService, ILogger<EventPublishingService> logger, M365AuthHelper auth
        , string mailBoxUser, string defaultLocale, string tenantId, bool sensivityReplicationEnabled, bool pdfConversionEnabled)
        : base(logger, auth)
    {
        _roleSrv = bodyRoleService;
        _notificationsService = notificationsService;
        _mailBoxUser = mailBoxUser;
        _defaultLocale = defaultLocale;
        _taskService = tasksService;
        _insideService = insideService;
        _configBodyService = configBodyService;
        _tenantId = Guid.Parse(tenantId);
        _sensivityReplicationEnabled = sensivityReplicationEnabled;
        _pdfConversionEnabled = pdfConversionEnabled;
        _outlookSyncService = outlookSyncService;
    }

    public async Task prepareRequirements(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));
        using var graphHelper = new GraphHelper(bodyCtx);

        var eventDAL = new InConstructionEventsDALprovider();

        var (CurrentState, PublishedVersion) = await eventDAL.GetByIdAndVersionLabel(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.ToCaptureState?.VersionLabel ?? throw new Exception("Version label is required"));

        await EnsureEventUserGroupForEvent(bodyCtx, CurrentState.AsChildCT<MeetingEnEdicionDAO>());
        await _outlookSyncService.EnsureEventMeeting(bodyCtx, graphHelper, CurrentState.AsChildCT<MeetingEnEdicionDAO>(), eventChangeInfo.BodyId);
        if (RequiredTeamsOnlineMeetingCreation(CurrentState))
            await EnsureEventOnlineMeeting(bodyCtx, graphHelper, CurrentState.AsChildCT<MeetingEnEdicionDAO>());
    }

    // TODO: refactor this method
    public async Task EventCaptureToDocumentSetUpdates(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));
        using var graphHelper = new GraphHelper(bodyCtx);

        // if still in construction or no changes or is cancelled from construction, do nothing
        var isCancelledWithoutPublishing = eventChangeInfo.StatusHasChanged && eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InConstruction) && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Cancelled);
        var eventIsInConstruction = !eventChangeInfo.StatusHasChanged && eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InConstruction);
        if (eventIsInConstruction || isCancelledWithoutPublishing || !eventChangeInfo.AnyChange)
            return;

        var eventsInConstructionprovider = new InConstructionEventsDALprovider();

        await bodyCtx.Site.EnsurepropertiesAsync(p => p.Id);
        var siteId = bodyCtx.Site.Id.ToString();

        var (CurrentState, PublishedVersion) = await eventsInConstructionprovider.GetByIdAndVersionLabel(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.ToCaptureState?.VersionLabel ?? throw new Exception("Version label is required"));

        var constructionEventListItem = CurrentState.AsListItem();
        await constructionEventListItem.EnsurepropertiesAsync(li => li.Folder.Queryproperties(f => f.Name), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);
        var constructionListId = constructionEventListItem.ParentList.Id.ToString();
        var constructionEventFolderName = constructionEventListItem.Folder.Name;


        // create/update published event and ensure all needed
        var publishedEvent = await EnsurePublishedEvent(bodyCtx, CurrentState.AsChildCT<MeetingEnEdicionDAO>());
        var publishedEventListItem = publishedEvent.AsListItem();
        var publishingList = publishedEventListItem.ParentList;
        var publishedEventFolder = publishedEventListItem.Folder;
        var publishedEventFolderName = publishedEventListItem.Folder.Name;
        var publishedEventFolderServerRelativeUrl = publishedEventFolder.ServerRelativeUrl;

        await publishingList.EnsurepropertiesAsync(p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
        var publishingListId = publishedEventListItem.ParentList.Id.ToString();

        var allCapturedFiles = eventChangeInfo?.Documents?.Updated.Concat(eventChangeInfo.Documents.Added).Concat(eventChangeInfo.Documents.NotChanged).ToList() ?? [];
        var allFilesWithUnsupportedLabels = await graphHelper.GetFilesIdsWithUnsupportedLabels(siteId, constructionListId, _tenantId, allCapturedFiles.Select(f => f.ItemUniqueId!.Value).ToArray());

        // files with following extensions will be converted to pdf, others wil be copied as is
        // from: https://learn.microsoft.com/en-us/graph/api/driveitem-get-content-format?view=graph-rest-1.0&tabs=http#format-options
        // TODO: move this to api env configuration
        var supportedExtensions = new string[] { ".doc", ".docx", ".ppt", ".pptm", ".pptx", ".rtf" };
        var allFilesToPdf = allCapturedFiles.Where(file => _pdfConversionEnabled &&
            supportedExtensions.Contains(Path.GetExtension(file.FileName)!.ToLower()) &&
            !allFilesWithUnsupportedLabels.Contains(file.ItemUniqueId!.Value)).ToList();
        var allFilesToCopy = allCapturedFiles.Except(allFilesToPdf).ToList();
        //allCapturedFiles.Select(f => f.ItemUniqueId!.Value).ToArray()

        // upload all files to the event folder an map source guids to destination paths (needed for related docs)
        // skip files that are not in the capture
        var copyResult = await PnPContentHelpers.CopyOrMoveFolderContents(bodyCtx, constructionEventListItem, publishedEventListItem, isMove: false, shouldSkipDelegate: (f) => !allFilesToCopy.Any(af => af.ItemUniqueId == f.ItemUniqueId));

        var srcUniqueIdToDstUniqueId = copyResult.CopiedSrcToDstIdMapping;

        // check for documents in destination that must be deleted (not in capture or skiped)
        var docsInDestinationThatShouldExist = srcUniqueIdToDstUniqueId.Values.ToArray();
        var docsInDestinationToDelete = copyResult.AllDstFiles.Where(dstFile => !docsInDestinationThatShouldExist.Any(ddst => ddst == dstFile.ItemUniqueId)).ToArray();
        var deleteBatch = ctx.NewBatch();
        foreach (var file in docsInDestinationToDelete)
            await file.BaseListItem.DeleteBatchAsync(deleteBatch);
        await ctx.ExecuteAsync(deleteBatch);
        var provider = new EventInConstructionDocumentDALprovider();
        var documentDAOs = await provider.GetInConstructionDocument(bodyCtx, CurrentState.IdMeeting);

        var tenantLabels = await graphHelper.GetTenatSensitivityLabels();

        foreach (var toPdfFile in allFilesToPdf)
        {
            try
            {
                var docDAO = documentDAOs.Where(x => x.FileRef == toPdfFile.FileServerRelativePath).First();

                var toPdfFileRelativePath = Path.DirectorySeparatorChar + constructionEventFolderName.UriCombine(toPdfFile.FolderRelativeUrl!);
                var newDocSetRelativeUrl = Path.ChangeExtension(toPdfFile.FolderRelativeUrl, ".pdf")!;
                var newPath = publishedEventFolderName.UriCombine(newDocSetRelativeUrl);

                //Get sensitivity labels if apply
                SensitivityLabelAssignment? sLabel = null;
                try
                {
                    var sLabelsResult = await graphHelper.GetFileSensitivityLabelsByPath(siteId, constructionListId, toPdfFileRelativePath);
                    sLabel = sLabelsResult?.Labels?.FirstOrDefault(l => l?.TenantId == _tenantId.ToString());
                }
                catch (Exception ex)
                {
                    Log.LogInformation($"File {toPdfFile.FileServerRelativePath}: labels not found or have errors.");
                    sLabel = null;
                }

                //If file has a label, remove it
                if (sLabel != null)
                {
                    if (docDAO != null && tenantLabels?.Any() == true)
                    {
                        docDAO.SensitivityLabelIsSystem = true;
                        await docDAO.AsListItem().UpdateAsync();
                    }

                    await graphHelper.SetFileSensitivityLabelsByPath(siteId, constructionListId, toPdfFileRelativePath, "", "Auto publishing remove label");

                    //Check if sensitivity label was removed
                    ExtractSensitivityLabelsResult? sLabelRemoved = null;

                    do
                    {
                        sLabelRemoved = await graphHelper.GetFileSensitivityLabelsByPath(siteId, constructionListId, toPdfFileRelativePath);
                        if (sLabelRemoved != null && sLabelRemoved.Labels != null && sLabelRemoved.Labels.Any(l => l.TenantId == _tenantId.ToString()))
                            await Task.Delay(3500);
                    } while (sLabelRemoved != null && sLabelRemoved.Labels != null && sLabelRemoved.Labels.Any(l => l.TenantId == _tenantId.ToString()));

                    if (docDAO != null && tenantLabels?.Any() == true)
                    {
                        docDAO.SensitivityLabelExpected = sLabel.SensitivityLabelId!.ToString();
                        docDAO.SensitivityLabelprocessing = true;
                        await docDAO.AsListItem().UpdateAsync();
                    }
                }

                var fileStream = await graphHelper.GetFileAsPdfByItemId(bodyCtx.Site.Id.ToString(), constructionListId, toPdfFile.ItemUniqueId.ToString()!);
                if (fileStream == null)
                {
                    Log.LogWarning($"File {toPdfFile.FileServerRelativePath} not found in event. Stream is null.");
                    continue;
                }

                using var ms = new MemoryStream();
                await fileStream.CopyToAsync(ms);
                var driveItem = await graphHelper.UploadSmallFileTo(bodyCtx.Site.Id.ToString(), publishingListId, newPath, ms.ToArray());

                //Set again the testvious label to original and transformed document
                if (sLabel != null && tenantLabels?.Any() == true)
                {
                    await graphHelper.SetFileSensitivityLabelsByPath(siteId, constructionListId, toPdfFileRelativePath, sLabel.SensitivityLabelId!, "Auto publishing");
                    await graphHelper.SetFileSensitivityLabelsByPath(siteId, publishingListId, Path.DirectorySeparatorChar + newPath, sLabel.SensitivityLabelId!, "Auto publishing");
                }

                var regexResult = GuidRegEx().Match(driveItem.ETag!); // hack: etag should store shp unique id (performance - skip additional queries) see https://karinebosch.wordtestss.com/my-articles/improdving-performance-of-sharepoint-sites/part-13-etag-header/
                if (regexResult.Success)
                    srcUniqueIdToDstUniqueId.Add(toPdfFile.ItemUniqueId!.Value, Guid.Parse(regexResult.Value));
                else
                    throw new Exception("Wrong SharePoint unique id handling");


            }
            catch (Exception ex)
            {
                Log.LogWarning($"File {toPdfFile.FileServerRelativePath} file is corrupt. Exception: {ex}");
            }
        }

        // update all related docs ids in agenda items and agreements
        await UpdateAllRelatedDocuments(Log, bodyCtx, srcUniqueIdToDstUniqueId, Source.Published, publishedEvent.IdMeeting);

        if (_sensivityReplicationEnabled)
            await ReplicateSensivityLabels(bodyCtx, graphHelper, _tenantId, constructionListId, publishingListId, allCapturedFiles, srcUniqueIdToDstUniqueId);

        await PnPContentHelpers.CaptureDocumentSet(bodyCtx, publishingList, publishedEventListItem, "Auto publishing");

        static async Task ReplicateSensivityLabels(IPnPContext bodyCtx, GraphHelper graphHelper, Guid tenantId, string constructionListId, string publishingListId, List<DocSetDocumentInformation> allCapturedFiles, Dictionary<Guid, Guid> srcUniqueIdToDstUniqueId)
        {
            await bodyCtx.Site.EnsurepropertiesAsync(p => p.SensitivityLabelId);
            // TODO: change site default sensivity label to library sensivity label
            Guid? defaultSensitiveLabelId = bodyCtx.Site.SensitivityLabelId.Equals(Guid.ECNTy) ? null : bodyCtx.Site.SensitivityLabelId;

            // get sensivity labels for all files
            var srcUniqueIdToLabelsMapping = await graphHelper.GetFilesSensitivityLabelIds(bodyCtx.Site.Id.ToString(), constructionListId, tenantId, allCapturedFiles.Select(f => f.ItemUniqueId!.Value).ToArray());

            // set sensivity labels for all files
            // select files with labels that are not site default and not eCNTy
            var docsToUpdateLabelInDest = srcUniqueIdToLabelsMapping
                .Where(pair => pair.Value.HasValue && (!defaultSensitiveLabelId.HasValue || defaultSensitiveLabelId.Value != pair.Value.Value))
                .Select(pair => (pair.Key, pair.Value!.Value));

            // using src to dst unique id mapping, set labels for files in destination
            var labelsToSetInDest = new Dictionary<Guid, Guid>();
            foreach (var (fileUniqueId, labelId) in docsToUpdateLabelInDest)
                if (srcUniqueIdToDstUniqueId.TryGetValue(fileUniqueId, out var fileInDestination))
                    labelsToSetInDest.Add(fileInDestination, labelId);

            await graphHelper.SetFilesSensitivityLabels(bodyCtx.Site.Id.ToString(), publishingListId, labelsToSetInDest, "Auto publishing");
        }
    }

    public async Task EventCaptureToTasksUpdates(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));

        // TODO: when task service is ready, replace all this code with a single call to task service
        // create/update attendance tasks for all users and manage their document access group
        if (eventChangeInfo.AttendeesHasChanged || eventChangeInfo.DetailsHasChanged)
        {
            await SyncEventAttendeesToAttendanceAndDelegationTasks(ctx, eventChangeInfo.BodyId, eventChangeInfo.SharedEventId, eventChangeInfo.AttendeesMustReconfirm);
        }

        try
        {
            var eventIsCancelled = eventChangeInfo.StatusHasChanged && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Cancelled);
            if (eventIsCancelled)
                await _taskService.CancelOrExpiredTasksBySharedEventId(bodyCtx, eventChangeInfo.BodyId, eventChangeInfo.SharedEventId, TaxonomyValuesIds.RequestStatus.Cancelled);
        }
        catch (Exception ex)
        {
            Log.LogError(ex, $"Error cancelling tasks for event {eventChangeInfo.SharedEventId}.");
        }

        try
        {
            var eventIsFinished = eventChangeInfo.StatusHasChanged && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Celebrated);
            if (eventIsFinished)
            {
                await _taskService.CancelOrExpiredTasksBySharedEventId(bodyCtx, eventChangeInfo.BodyId, eventChangeInfo.SharedEventId, TaxonomyValuesIds.RequestStatus.Expired);
            }

        }
        catch (Exception ex)
        {
            Log.LogError(ex, $"Error exceeding tasks for event {eventChangeInfo.SharedEventId}.");
        }
    }

    public async Task EventCaptureToNotifications(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        var firstPublishOfEvent = eventChangeInfo.FromCaptureState == null;

        if (eventChangeInfo.SkipUserCommunications)
            return;

        var isEventCancelledFromInConstruction = eventChangeInfo.StatusHasChanged &&
                                                 eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Cancelled) &&
                                                 eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InConstruction);

        if (isEventCancelledFromInConstruction)
            return; // no notification are needed

        var notificationsToSend = GetNotificationsByStatus(eventChangeInfo);

        if (notificationsToSend?.Values.Count > 0)
        {
            var bodiesToNotify = notificationsToSend.SelectMany(pair => pair.Value).Distinct().ToList();
            var bodiesToObtain = bodiesToNotify.Where(v => !v.Equals(BodyRole.AllInvitedUsers)).ToList();

            var notificationInfo = new Dictionary<BodyRole, IEnumerable<string>?>();

            if (bodiesToObtain.Count > 0)
            {
                var bodiesTasks = new List<Task<IEnumerable<UserBodyRoles>>>();
                foreach (var body in bodiesToObtain)
                {
                    bodiesTasks.Add(_roleSrv.GetBodyUsersById(eventChangeInfo.BodyId, body));
                    notificationInfo.Add(body, null);
                }

                Task.WaitAll([.. bodiesTasks]);

                for (var i = 0; i < bodiesTasks.Count; i++)
                {
                    var result = bodiesTasks[i].Result;
                    notificationInfo[bodiesToObtain[i]] = result.Select(r => r.UserPrincipalName).ToList();
                }
            }

            if (bodiesToNotify.Contains(BodyRole.AllInvitedUsers))
            {
                notificationInfo.Add(BodyRole.AllInvitedUsers, eventChangeInfo.AllAttendees);
            }

            var ProfileService = new ProfileService();
            // userslocale tiene email - idioma elegido.
            // var usersLocale = await ProfileService.GetUserstestferredLanguage(ctx, notificationInfo.Values.SelectMany(v => v).Select(v => v).Distinct().ToList());
            var url = _roleSrv.GetNotificationUrlByEvent(ctx, eventChangeInfo.SharedEventId, eventChangeInfo.BodyId);
            var usersInBody = new List<string>();

            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));
            var eventsprovider = new InConstructionEventsDALprovider();

            if (!firstPublishOfEvent)
            {
                var (currentVersion, publishedVersion, changedFields) = await eventsprovider.GetChangesFromVersion(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.FromCaptureState!.VersionId);
                var pendingChanges = GetPendingChanges(eventChangeInfo, changedFields, currentVersion, publishedVersion);

                //AMC: AÑADIMOS POR DEFECTO EL NOMBRE DEL ÓRGANO, QUE EN ALGUNAS PLANTILLAS ESTÁ
                //notificationValues.Add(ReplacedTokens.DepartmentName, currentVersion.Department?.Label ?? string.ECNTy);
                var notificationValues = GetContosoEmailValuesToReplace(bodyCtx, currentVersion, eventChangeInfo);

                foreach (var notificationToSend in notificationsToSend)
                {
                    foreach (var body in notificationToSend.Value)
                    {
                        usersInBody.AddRange(notificationInfo[body] ?? []);
                    }
                    try
                    {
                        bool includePendingChanges = notificationToSend.Key == NotificationsMessages.Meeting.Modificar || notificationToSend.Key == NotificationsMessages.Meeting.ModificarOrdenDia || notificationToSend.Key == NotificationsMessages.Meeting.ModificarCertificados;
                        await _notificationsService.AddUserNotification(eventChangeInfo.BodyId, publishedVersion.MeetingType!.TermId, usersInBody.Distinct().ToList(), notificationToSend.Key, this._defaultLocale, url, notificationValues, includePendingChanges ? pendingChanges : []);
                    }
                    catch (Exception ex)
                    {
                        Log.LogError(ex, $"Error sending notification {notificationToSend.Key} for event {eventChangeInfo.SharedEventId}.");
                    }
                }
            }
        }
    }

    public async Task SyncEventAgendaItemsAndAgreements(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
        var eventDAL = new InConstructionEventsDALprovider();

        var constructionEvent = await eventDAL.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Error sync agenda, event '{sharedEventId}' not found in body '{bodyId}'");
        var constructionEventListItem = constructionEvent.AsListItem();

        var publishedEvent = await EnsurePublishedEvent(bodyCtx, constructionEvent.AsChildCT<MeetingEnEdicionDAO>());
        var publishedEventListItem = publishedEvent.AsListItem();

        var copyResult = await PnPContentHelpers.CopyOrMoveFolderContents(bodyCtx, constructionEventListItem, publishedEventListItem, isMove: false, shouldSkipDelegate: (f) =>
            !f.ContentTypeId.StartsWith(ContentTypeIds.AgendaItem)
        );
    }

    public async Task SyncEventAttendeesToAttendanceAndDelegationTasks(IPnPContext ctx, string bodyId, string sharedEventId, bool attendeesMustReconfirm)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
        var eventDAL = new InConstructionEventsDALprovider();
        var attendanceDAL = new EventAttendanceRequestDALprovider(eventDAL);
        var userTasksprovider = new TasksDALprovider();

        var (theEvent, existingAttendancesRequests) = await attendanceDAL.GetAllDTOByEvent(bodyCtx, sharedEventId);
        var storedEvent = theEvent.AsChildCT<MeetingEnEdicionDAO>();
        var bodyUsers = await _roleSrv.GetBodyUsersById(bodyId);
        var eventGroup = await EnsureEventUserGroupForEvent(bodyCtx, storedEvent);

        var eventHasStarted = storedEvent.EstadoMeeting!.TermId.Equals(TaxonomyValuesIds.EventStatus.InCelebration) || storedEvent.EstadoMeeting!.TermId.Equals(TaxonomyValuesIds.EventStatus.Celebrated);
        var statusForNewAttendance = eventHasStarted ? TaxonomyValuesIds.RequestStatus.Expired : TaxonomyValuesIds.RequestStatus.Pending;
        var statusForDelegatedTasks = eventHasStarted ? TaxonomyValuesIds.RequestStatus.Expired : TaxonomyValuesIds.RequestStatus.Pending;

        var (added, removed, keeped, notFound) = await bodyCtx.SetSPGroupMembershipTo(eventGroup.Id, storedEvent.Asistentes.AsUserPrincipalName());

        await bodyCtx.Web.LoadAsync(p => p.AssociatedOwnerGroup);

        var allUsers = added.Concat(keeped).ToList();

        // create collection for new, updated and removed attendances
        var toRemoveAttendancesRequests = existingAttendancesRequests.Where((ea) => !allUsers.Any((u) => u.Id == ea?.AssignedTo.FirstOrDefault()?.LookupId));
        var toUpdateAttendancesRequests = existingAttendancesRequests.Where((ea) => allUsers.Concat(keeped).ToList().Any((u) => u.Id == ea?.AssignedTo.FirstOrDefault()?.LookupId) || allUsers.Concat(added).ToList().Any((u) => u.Id == ea?.AssignedTo.FirstOrDefault()?.LookupId));

        //var toUpdateAttendancesRequests = toUpdateAttendancesRequests + allUsers.Where((eaId) => existingAttendancesRequests.Where((ea) => eaId.Id == ea?.AssignedTo.FirstOrDefault()?.LookupId));
        var toAddAttendancesRequests = allUsers
            .Where((eaId) => !existingAttendancesRequests.Any((ear) => eaId.Id == ear?.AssignedTo.FirstOrDefault()?.LookupId))
            .Select((a) => new EventAttendanceDTO()
            {
                //TODO revisar si pueden ser grupos o personas
                Title = storedEvent.Title,
                AssignedTo = [new(a)],
                RequestedBy = new FieldUserValue(bodyCtx.Web.AssociatedOwnerGroup),
                RequestStatus = statusForNewAttendance,
                AttendanceTypeId = storedEvent.MeetingType?.TermId.ToString(),
                CanVote = bodyUsers.Any(bu => bu.UserPrincipalName.Equals(a.UserPrincipalName, StringComparison.OrdinalIgnoreCase)
                    && bu.UserRoles.HasFlag(BodyRole.Member)),
            }) ?? [];

        //update existing attendances: cancel removed users, update attendance type and request status if needed
        var delgationTasksToUpdate = new List<TasksDTO>();
        foreach (var attRequest in toRemoveAttendancesRequests)
        {
            attRequest.RequestStatus = TaxonomyValuesIds.RequestStatus.Cancelled; // users removed

            // TODO: get all and filter by user in memory
            var delegatedTasks = await userTasksprovider.GetDelegatedTaskByIdMeetingAndIdUserDelegated(bodyCtx, sharedEventId, attRequest.AssignedTo.First().LookupId);

            if (!(delegatedTasks?.Count() > 0))
                continue;

            delgationTasksToUpdate.AddRange(delegatedTasks.Select(dtask =>
            {
                dtask.EstadoRequest = statusForDelegatedTasks;
                return dtask;
            }).ToList());
        }
        await userTasksprovider.AddOrUpdateForEvent(bodyCtx, sharedEventId, delgationTasksToUpdate);

        foreach (var attRequest in toUpdateAttendancesRequests)
        {
            attRequest.RequestStatus = attendeesMustReconfirm && !eventHasStarted ? TaxonomyValuesIds.RequestStatus.Pending : (attRequest.RequestStatus == TaxonomyValuesIds.RequestStatus.Cancelled ? TaxonomyValuesIds.RequestStatus.Pending : attRequest.RequestStatus);
            attRequest.AttendanceTypeId = attRequest.RequestStatus == TaxonomyValuesIds.RequestStatus.Pending ? storedEvent.MeetingType?.TermId.ToString() ?? string.ECNTy : (attendeesMustReconfirm && !eventHasStarted ? storedEvent.MeetingType?.TermId.ToString() ?? string.ECNTy : attRequest.AttendanceTypeId ?? string.ECNTy);
        }

        await attendanceDAL.AddOrUpdateItemsForEvent(bodyCtx, storedEvent.IdMeeting, [.. toRemoveAttendancesRequests, .. toUpdateAttendancesRequests, .. toAddAttendancesRequests]);
    }

    private static List<PendingChanges> GetPendingChanges(EventSendOperationInformation eventChangeInfo, string[] changedFields, BaseEventDAO currentVersion, BaseEventDAO publishedVersion)
    {
        var pendingChanges = new List<PendingChanges>();
        var pendingChangesDALprovider = new EventPendingChangeDALprovider();
        if (eventChangeInfo.AttendeesHasChanged)
        {
            var attendanceChanges = FilterAddOrRemovedAttendance(eventChangeInfo.Attendees);
            var pendingChangesAttendance = attendanceChanges is not null ? EventPendingChangeDALprovider.FilterAttendance(attendanceChanges) : [];
            pendingChanges.AddRange(pendingChangesAttendance);
        }

        if (eventChangeInfo.OrderOfDayHasChanged || eventChangeInfo.ModifiedOrderOfDayName is not null || eventChangeInfo.DocumentsHasChanged)
        {
            var pendingChangesDocs = eventChangeInfo.ChangedDocuments is not null ? EventPendingChangeDALprovider.FilterAddRemoveUpdateDocs(eventChangeInfo.ChangedDocuments) : [];
            pendingChanges.AddRange(pendingChangesDocs);
        }

        if (eventChangeInfo.DetailsHasChanged)
        {
            var eventChanges = pendingChangesDALprovider.FilterEventChanges(currentVersion, changedFields, publishedVersion);
            var pendingChangesEvent = eventChanges is not null ? pendingChangesDALprovider.processEventChanges(eventChanges) : [];
            pendingChanges.AddRange(pendingChangesEvent);

        }
        return pendingChanges;
    }

    public static List<AttendanceChanges> FilterAddOrRemovedAttendance(AttendeesChangesResult? attendees)
    {
        var attendanceChanges = new List<AttendanceChanges>();

        if (attendees != null)
        {
            attendanceChanges.AddRange(attendees.Added.Select(user => new AttendanceChanges { AddOrRemoved = "AddAttendance", user = user }));
            attendanceChanges.AddRange(attendees.Removed.Select(user => new AttendanceChanges { AddOrRemoved = "RemoveAttendance", user = user }));
        }

        return attendanceChanges;
    }

    public async Task EventCaptureToTeamsAndOutlookUpdates(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));

        var eventDAL = new InConstructionEventsDALprovider();
        var (CurrentState, PublishedVersion) = await eventDAL.GetByIdAndVersionLabel(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.ToCaptureState?.VersionLabel ?? throw new Exception("Version label is required"));

        var eventIsInConstruction = !eventChangeInfo.StatusHasChanged && eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InConstruction);
        if (eventIsInConstruction)
            return;

        var currentState = CurrentState.AsChildCT<MeetingEnEdicionDAO>();

        if (!eventChangeInfo.SkipUserCommunications)
        {
            await _outlookSyncService.UpdateMeetingEvent(bodyCtx, eventChangeInfo, PublishedVersion, currentState);

            //Send notification by mail
            if (new[] { TaxonomyValuesIds.EventStatus.Published, TaxonomyValuesIds.EventStatus.Reserved }
            .Contains(string.IsNullOrECNTy(eventChangeInfo.NewEventStatusId) ? eventChangeInfo.CurrentEventStatusId : eventChangeInfo.NewEventStatusId))
            {
                List<string> newAttendees = [];
                var (currentVersion, publishedVersion, changedFields) = (new BaseEventDAO(), new BaseEventDAO(), Array.ECNTy<string>());
                if (eventChangeInfo.FromCaptureState != null)
                {
                    //If it is not a new event, only new attendees are selected
                    (currentVersion, publishedVersion, changedFields) = await eventDAL.GetChangesFromVersion(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.FromCaptureState!.VersionId);
                    var pendingChanges = GetPendingChanges(eventChangeInfo, changedFields, currentVersion, publishedVersion);
                    newAttendees = pendingChanges.Where(x => x.ChangeType == "AddAttendance").Select(x => x.ChangeValue).ToList();
                }
                else
                {
                    //If it is a new event, all the attendees are selected
                    newAttendees = eventChangeInfo.AllAttendees;
                    currentVersion = currentState;
                }

                var recipients = _outlookSyncService.GetRecipients(eventChangeInfo, newAttendees);
                var octestcipients = await _outlookSyncService.GetOCtestcipients(bodyCtx, eventChangeInfo);
                octestcipients = octestcipients.Except(recipients).ToList();
                await _outlookSyncService.SendMeetingEventByEmail(bodyCtx, eventChangeInfo, currentVersion, recipients, octestcipients);
            }
        }


        // create online teams meeting if needed
        await UpdateOnlineMeeting(bodyCtx, eventChangeInfo, currentState);
    }

    public async Task EventCaptureToRemovalOrArchive(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));
        await bodyCtx.Site.EnsurepropertiesAsync(s => s.ServerRelativeUrl, s => s.Id);
        using var scope = Log.BeginContosoScope("EventCaptureToRemovalOrArchive", eventChangeInfo.BodyId, eventChangeInfo.SharedEventId, eventChangeInfo.SentBy);

        using var graphHelper = new GraphHelper(bodyCtx);
        var inConstrucctionEventsprovider = new InConstructionEventsDALprovider();
        var publishedEventsprovider = new PublishedEventsDALprovider();
        var archivedEventsprovider = new ArchivedEventsDALprovider();
        var attendanceDAL = new EventAttendanceRequestDALprovider(inConstrucctionEventsprovider);
        var (CurrentState, PublishedVersion) = await inConstrucctionEventsprovider.GetByIdAndVersionLabel(bodyCtx, eventChangeInfo.EventListItemId, eventChangeInfo.ToCaptureState?.VersionLabel ?? throw new Exception("Version label is required"));

        var eventIsCancelled = eventChangeInfo.StatusHasChanged && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Cancelled);
        var eventIsArchived = eventChangeInfo.StatusHasChanged && eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.Archived);

        Log.LogDebug($"Event (status changed: {eventChangeInfo.StatusHasChanged}, status: {eventChangeInfo.CurrentEventStatusId}, new status: {eventChangeInfo.NewEventStatusId}): Cancelled: {eventIsCancelled} Archived: {eventIsArchived}");

        // event will be deleted only on status change from InConstruction to Cancelled
        if (eventIsCancelled)
        {
            if (eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InConstruction))
            {
                await inConstrucctionEventsprovider.Delete(bodyCtx, CurrentState.ID);
                Log.LogInformation("Event was cancelled and has been deleted as testvious status was in construction");
            }
            else
            {
                Log.LogInformation("Published event was cancelled"); // TODO: check if archival is needed.
            }
        }

        if (eventIsArchived)
        {
            try
            {
                // ensure that all finished documents has eni xml and index is created
                var publishedEvent = await EnsureEniXmlAndIndexForPublishedEventDocuments(ctx, eventChangeInfo);

                // move published event and all its contents to archived library
                var destinationAbsoluteUrl = "/".UriCombine(new Uri(bodyCtx.Uri.ToString().UriCombine(ListsSiteRelativeUrls.ArchivedEvents)).GetComponents(UriComponents.Path, UriFormat.Unescaped));

                Log.LogDebug($"Moving published event to {destinationAbsoluteUrl}");
                var moveResult = await PnPContentHelpers.CopyOrMoveFolder(bodyCtx, publishedEvent.AsListItem(), destinationAbsoluteUrl, new CopyMigrationOptions
                {
                    AllowSchemaMismatch = true,
                    AllowSmallerVersionLimitOnDestination = true,
                    IgnoreVersionHistory = false,
                    IsMoveMode = true,
                    BypassSharedLock = true,
                    ExcludeChildren = false,
                    NameConflictBehavior = SPMigrationNameConflictBehavior.Replace // should not happen
                });

                Log.LogDebug($"Saving construction data of event to {destinationAbsoluteUrl}");
                // move tasks (or potentially any document) from construction to archived
                await SaveInConstructionData(bodyCtx, CurrentState, publishedEvent, destinationAbsoluteUrl);

                // ensure archived docset data
                var archivedEvent = (await archivedEventsprovider.GetBySharedEventId(bodyCtx, CurrentState.IdMeeting))!;
                var archivedEventListItem = archivedEvent.AsListItem();
                await archivedEventListItem.EnsurepropertiesAsync(li => li.UniqueId, li => li.Folder.Queryproperties(f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(pl => pl.Id, pl => pl.RootFolder.Queryproperties(rf => rf.ServerRelativeUrl), p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title)));
                archivedEvent.ContentTypeId = await archivedEventsprovider.GetContentTypeIdValueInList(bodyCtx);
                await archivedEvent.AsListItem().UpdateAsync();

                Log.LogDebug($"Updating related documents of event");
                await UpdateAllRelatedDocuments(Log, bodyCtx, moveResult.CopiedSrcToDstIdMapping, Source.Archived, archivedEvent.IdMeeting);

                Log.LogDebug($"Capturing archived doc set of event");
                await graphHelper.CaptureDocumentSet(bodyCtx.Site.Id.ToString(), archivedEventListItem.ParentList.Id.ToString(), archivedEventListItem.UniqueId.ToString(), false, "Archived event");

                // remove in construction data (recicle)
                await inConstrucctionEventsprovider.Delete(bodyCtx, CurrentState.ID);
                Log.LogInformation("Event was archived and has been deleted from in construction");
            }
            catch (Exception ex)
            {
                Log.LogError(ex, $"Error event capture to archive for body  {eventChangeInfo.BodyId}.");
            }
        }

        static async Task SaveInConstructionData(IPnPContext bodyCtx, BaseEventDAO inConstructionEvent, BaseEventDAO publishedEvent, string destinationAbsoluteUrl)
        {
            var archivedDocSetAbsolutePath = destinationAbsoluteUrl.UriCombine(Path.GetFileName(publishedEvent.FileRef));
            var archivedDocSetServerRelativePath = bodyCtx.Site.ServerRelativeUrl.ToString().UriCombine(ListsSiteRelativeUrls.ArchivedEvents).UriCombine(Path.GetFileName(publishedEvent.FileRef));

            var inConstructionEventListItem = inConstructionEvent.AsListItem();
            await inConstructionEventListItem.EnsurepropertiesAsync(li => li.UniqueId, li => li.Folder.Queryproperties(f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(pl => pl.Id, pl => pl.RootFolder.Queryproperties(rf => rf.ServerRelativeUrl), p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title)));

            var inConstructionList = inConstructionEventListItem.ParentList;
            var inConstructionEventFolder = inConstructionEventListItem.Folder;

            // get all files in the event folder
            var listOfFiles = await inConstructionList.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
            {
                ViewXml = PnPContentHelpers.CamlViewBuilder("<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>", [Fields.UniqueId, nameof(BaseSPListItemDAO.FileRef), nameof(BaseSPListItemDAO.ContentTypeId), nameof(BaseSPListItemDAO.Created)], null, inConstructionEventFolder.ServerRelativeUrl, ViewScope.Recursive),
                FolderServerRelativeUrl = inConstructionEventFolder.ServerRelativeUrl
            }, (item) => new
            {
                UniqueId = Guid.Parse((string)item.Values[Fields.UniqueId]),
                ServerRelativeUrl = (string)item.Values[nameof(BaseSPListItemDAO.FileRef)],
                ContentTypeId = (string)item.Values[nameof(BaseSPListItemDAO.ContentTypeId)],
                Created = (DateTime)item.Values[nameof(BaseSPListItemDAO.Created)]
            });

            // keep all old tasks files (attendance is required for attendance tab)
            listOfFiles = listOfFiles
                .Where(f => f.ContentTypeId.StartsWith(ContentTypeIds.Task.Request))
                .ToList();

            var filePathsToMove = new List<string>();
            foreach (var file in listOfFiles)
                filePathsToMove.Add(bodyCtx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(file.ServerRelativeUrl));

            if (filePathsToMove.Count != 0)
            {
                await bodyCtx.Site.EnsureCopyJobHasFinishedAsync(await bodyCtx.Site.CreateCopyJobsAsync(
                    [.. filePathsToMove]
                    , bodyCtx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(archivedDocSetAbsolutePath.UriCombine(HiddenFolderName)),
                    new CopyMigrationOptions
                    {
                        AllowSchemaMismatch = true,
                        AllowSmallerVersionLimitOnDestination = true,
                        IgnoreVersionHistory = false,
                        IsMoveMode = true,
                        BypassSharedLock = true,
                        ExcludeChildren = false,
                        NameConflictBehavior = SPMigrationNameConflictBehavior.Replace // should not happen
                    }));
            }
        }
    }

    public async Task<BaseEventDAO> EnsureEniXmlAndIndexForPublishedEventDocuments(IPnPContext ctx, EventSendOperationInformation eventChangeInfo)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(eventChangeInfo.BodyId));
        using var graphHelper = new GraphHelper(bodyCtx);
        var eniExtensionFile = ".eni.xml";
        var eniIndexNameFile = "index";

        // get configuration or fail if not correct
        var (dir, sia) = await _configBodyService.GetFieldsSiaDir3ConfigDepartments(eventChangeInfo.BodyId);

        // use for testing:
        //sia = "3033285";
        //dir = "EA0020003";

        if (string.IsNullOrECNTy(dir) || string.IsNullOrECNTy(sia)) throw new Exception($"Body {eventChangeInfo.BodyId} is missing dir3 or sia values in configuration");

        var publishedEventsprovider = new PublishedEventsDALprovider();
        var publishedEvent = await publishedEventsprovider.GetBySharedEventId(bodyCtx, eventChangeInfo.SharedEventId) ?? throw new InvalidOperationException($"Published event not found for event {eventChangeInfo.SharedEventId}.");

        var publishedEventListItem = publishedEvent.AsListItem();
        await publishedEventListItem.EnsurepropertiesAsync(li => li.UniqueId, li => li.Folder.Queryproperties(f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(pl => pl.Id, pl => pl.RootFolder.Queryproperties(rf => rf.ServerRelativeUrl), p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title)));

        var publishedList = publishedEventListItem.ParentList;
        var publishedEventFolder = publishedEventListItem.Folder;

        // get all files in the event folder
        var listOfFiles = await publishedList.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
        {
            ViewXml = PnPContentHelpers.CamlViewBuilder("<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>", [Fields.UniqueId, nameof(BaseSPListItemDAO.FileRef), nameof(BaseSPListItemDAO.ContentTypeId), nameof(BaseSPListItemDAO.Created)], null, publishedEventFolder.ServerRelativeUrl, ViewScope.Recursive),
            FolderServerRelativeUrl = publishedEventFolder.ServerRelativeUrl
        }, (item) => new
        {
            UniqueId = Guid.Parse((string)item.Values[Fields.UniqueId]),
            ServerRelativeUrl = (string)item.Values[nameof(BaseSPListItemDAO.FileRef)],
            ContentTypeId = (string)item.Values[nameof(BaseSPListItemDAO.ContentTypeId)],
            Created = (DateTime)item.Values[nameof(BaseSPListItemDAO.Created)]
        });

        // remove hidden files and existing eni documents
        listOfFiles = listOfFiles
            .Where(f => !(f.ServerRelativeUrl.Contains(HiddenFolderName, StringComparison.OrdinalIgnoreCase) || f.ContentTypeId.StartsWith(ContentTypeIds.DocumentEniPTCAAPP)))
            .ToList();

        var filesCount = listOfFiles.Count();
        var eniFileCollection = new List<InputDossierFile>(filesCount);
        var eniFileUpdateRequests = new List<GraphHelper.DriveItemsFieldsUpdateRequest>(filesCount);

        Log.LogDebug($"Files to create ENI: {filesCount}. When 0, no index is created.");

        if (!listOfFiles.Any()) return publishedEvent; // if no documents to archive do not create index.

        foreach (var file in listOfFiles)
        {
            var type = file.ContentTypeId switch
            {
                var id when id.StartsWith(ContentTypeIds.DocumentCertificacion) => DissierFileType.Certificate,
                var id when id.StartsWith(ContentTypeIds.MinutesDocument) => DissierFileType.Act,
                _ => DissierFileType.Other
            };

            var fileContent = await FileHelper.DownloadFileAsByteArray(bodyCtx, file.ServerRelativeUrl);
            var eniFile = new InputDossierFile() { Content = fileContent, CaptureDate = file.Created, UniqueId = file.UniqueId, DIR3 = [dir], DocumentType = type };
            var eniFileContent = await _insideService.DocumentToENIAsync(eniFile);

            var eniDriveRelativePath = file.ServerRelativeUrl.Replace(publishedList.RootFolder.ServerRelativeUrl, "") + eniExtensionFile;

            var driveItem = await graphHelper.UploadSmallFileTo(bodyCtx.Site.Id.ToString(), publishedList.Id.ToString(), eniDriveRelativePath, eniFileContent);

            eniFileCollection.Add(eniFile);
            eniFileUpdateRequests.Add(new GraphHelper.DriveItemsFieldsUpdateRequest()
            {
                ContentTypeId = ContentTypeIds.DocumentEniPTCAAPP,
                DriveId = driveItem.ParentReference!.DriveId!,
                DriveItemId = driveItem.Id!,
                FieldValues = new Dictionary<string, object>() { { Fields.OriginalDocumentId, file.UniqueId.ToString() } }
            });

            Log.LogDebug($"ENI file uploaded: {eniDriveRelativePath} (original document: {file.UniqueId} DIR3: {dir}, type: {type})");
        }

        var index = new InputDossier()
        {
            CaptureDate = publishedEvent.Created,
            UniqueId = publishedEvent.AsListItem().UniqueId,
            Files = [.. eniFileCollection],
            ClasificationSIA = sia,
            RootFolderName = "Contoso",
            FooterText = "Contoso",
            BodyNames = ["Contoso"]
        };
        var eniIndexFileContent = await _insideService.CreateIndexAsync(index);
        var eniIndexDriveRelativePath = publishedEventFolder.ServerRelativeUrl.Replace(publishedList.RootFolder.ServerRelativeUrl, "").UriCombine(eniIndexNameFile + eniExtensionFile);
        var indexDriveItem = await graphHelper.UploadSmallFileTo(bodyCtx.Site.Id.ToString(), publishedList.Id.ToString(), eniIndexDriveRelativePath, eniIndexFileContent);
        eniFileUpdateRequests.Add(new GraphHelper.DriveItemsFieldsUpdateRequest()
        {
            ContentTypeId = ContentTypeIds.DocumentEniPTCAAPP,
            DriveId = indexDriveItem.ParentReference!.DriveId!,
            DriveItemId = indexDriveItem.Id!,
        });

        await graphHelper.UpdateListItemFieldsInBatch(eniFileUpdateRequests);

        Log.LogInformation($"Dossier was created (${eniIndexDriveRelativePath}) with ${filesCount} ENI files");

        return publishedEvent;
    }

    private async Task<OnlineMeeting> EnsureEventOnlineMeeting(IPnPContext bodyCtx, GraphHelper graphHelper, MeetingEnEdicionDAO storedEvent)
    {
        OnlineMeeting onlineMeeting;

        if (string.IsNullOrECNTy(storedEvent.IdReunionOnlineMSTeams))
        {
            onlineMeeting = await graphHelper.CreateOnlineMeeting(_mailBoxUser, new OnlineMeeting()
            {
                Subject = storedEvent.Title,
                // RecordAutomatically = true,
                LobbyBypassSettings = new LobbyBypassSettings() { Scope = LobbyBypassScope.Invited },
                IsEntryExitAnnounced = true,
                Allowedtestsenters = OnlineMeetingtestsenters.RoleIstestsenter,
                AllowMeetingChat = MeetingChatMode.Enabled,
                ShareMeetingChatHistoryDefault = MeetingChatHistoryDefaultMode.All,
                // Watermarkprodtection = new Microsoft.Graph.Models.WatermarkprodtectionValues() { IsEnabledForVideo = true, IsEnabledForContentSharing = true }, // https://learn.microsoft.com/en-us/MicrosoftTeams/watermark-meeting-content-video
                // AllowTeamworkReactions = false,
                JoinMeetingIdSettings = new JoinMeetingIdSettings() { IsPasscodeRequired = false }
            });

            var dal = new InConstructionEventsDALprovider();
            var freshItem = await dal.GetById(bodyCtx, storedEvent.ID) ?? throw new Exception($"Event with id {storedEvent.ID} was not found"); ;
            freshItem.IdReunionOnlineMSTeams = onlineMeeting.Id!;
            freshItem.UrlOnlineTool = onlineMeeting.JoinWebUrl.AsUrlFieldValue();
            await freshItem.AsListItem().SystemUpdateAsync();
        }
        else
        {
            onlineMeeting = await graphHelper.GetOnlineMeetingById(_mailBoxUser, storedEvent.IdReunionOnlineMSTeams);
        }

        return onlineMeeting;
    }

    private async Task<ISharePointGroup> EnsureEventUserGroupForEvent(IPnPContext bodyCtx, MeetingEnEdicionDAO storedEvent)
    {
        ISharePointGroup group;
        var storedGroupId = storedEvent.GrupoAsistentesAsociado?.LookupId ?? -1;

        if (storedGroupId == -1)
        {
            var attendanceGroupName = EventGroupprefix + storedEvent.Title.GetSafeName().TruncateString(80) + " (" + storedEvent.ID + ")";
            // check if group already exists (should not)
            group = await bodyCtx.Web.SiteGroups.FirstOrDefaultAsync(g => g.Title == attendanceGroupName);
            if (group == null)
            {
                // create event attendance security group
                await bodyCtx.Web.LoadAsync(p => p.AssociatedOwnerGroup);
                group = await bodyCtx.Web.SiteGroups.AddAsync(attendanceGroupName);
                await group.SetUserAsOwnerAsync(bodyCtx.Web.AssociatedOwnerGroup.Id);
                await group.UpdateAsync();
            }

            // update the event with the group
            var dal = new InConstructionEventsDALprovider();
            var freshItem = await dal.GetById(bodyCtx, storedEvent.ID) ?? throw new Exception($"Event with id {storedEvent.ID} was not found"); ;
            freshItem.GrupoAsistentesAsociado = new FieldUserValue(group);
            await freshItem.AsListItem().SystemUpdateAsync();

            Log.LogInformation($"Event SharePoint security group {attendanceGroupName} created for event {storedEvent.ID}");
        }
        else
            group = await bodyCtx.Web.SiteGroups.FirstOrDefaultAsync(g => g.Id == storedGroupId);

        return group;
    }

    private async Task<MeetingPublicadaDAO> EnsurePublishedEvent(IPnPContext bodyCtx, MeetingEnEdicionDAO inConstructionEvent)
    {
        var publishEventsprovider = new PublishedEventsDALprovider();

        var constructionEventListItem = inConstructionEvent.AsListItem();
        await constructionEventListItem.EnsurepropertiesAsync(li => li.Folder.Queryproperties(f => f.Name, f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);
        var docSetName = constructionEventListItem.Folder.Name;

        var publishedEvent = await publishEventsprovider.CreateOrUpdateFrom(bodyCtx, docSetName, inConstructionEvent.AsChildCT<MeetingEnEdicionDAO>());
        var publishedEventListItem = publishedEvent.AsListItem();
        await publishedEventListItem.EnsurepropertiesAsync(li => li.Folder.Queryproperties(f => f.Name, f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);

        var eventGroup = await EnsureEventUserGroupForEvent(bodyCtx, inConstructionEvent);
        await bodyCtx.Web.LoadAsync(p => p.AssociatedVisitorGroup, p => p.RoleDefinitions.Queryproperties(r => r.Id, r => r.Name));

        var publishedListItem = publishedEvent.AsListItem();

        await publishedListItem.BreakRoleInheritanceAsync(false, true);

        await publishedListItem.AddRoleDefinitionsAsync(eventGroup.Id, ["Leer"]);
        await publishedListItem.AddRoleDefinitionsAsync(bodyCtx.Web.AssociatedVisitorGroup.Id, ["Leer"]); // TODO: check who is who and shp user groups.

        Log.LogInformation($"Event published folder is now avalible");

        return publishedEvent;
    }

    private Dictionary<string, List<BodyRole>> GetNotificationsByStatus(EventSendOperationInformation eventChangeInfo)
    {
        var result = new Dictionary<string, List<BodyRole>>();

        if (eventChangeInfo.StatusHasChanged)
        {
            switch (eventChangeInfo.NewEventStatusId)
            {
                case TaxonomyValuesIds.EventStatus.InConstruction:
                    result.Add(NotificationsMessages.Meeting.Nueva, [BodyRole.Scheduler, BodyRole.SchedulerAssistant]);
                    break;
                case TaxonomyValuesIds.EventStatus.Published:
                    result.Add(NotificationsMessages.Meeting.Publicar, [BodyRole.AllInvitedUsers]);
                    break;
                case TaxonomyValuesIds.EventStatus.InCelebration:
                    result.Add(NotificationsMessages.Meeting.EnCelebracion, [BodyRole.AllInvitedUsers]);
                    break;
                case TaxonomyValuesIds.EventStatus.Celebrated:
                    result.Add(NotificationsMessages.Meeting.Celebrada, [BodyRole.AllInvitedUsers]);
                    break;
                case TaxonomyValuesIds.EventStatus.Archived:
                    result.Add(NotificationsMessages.Meeting.Archivada, [BodyRole.Scheduler, BodyRole.SchedulerAssistant]);
                    break;
                case TaxonomyValuesIds.EventStatus.Cancelled:
                    result.Add(NotificationsMessages.Meeting.Cancelada, [BodyRole.AllInvitedUsers]);
                    break;
            }
        }

        if (result.Count == 0) //AGP si no es por cambio de estado, enviamos más Notifications
        {
            var detectedChange = false;

            if (eventChangeInfo.OrderOfDayHasChanged)
            {
                detectedChange = true;
                result.Add(NotificationsMessages.Meeting.ModificarOrdenDia, [BodyRole.AllInvitedUsers]);
            }

            if (eventChangeInfo.CertificateOfAgreementHasChanged)
            {
                detectedChange = true;
                result.Add(NotificationsMessages.Meeting.ModificarCertificados, [BodyRole.AllInvitedUsers, BodyRole.Member]);
            }

            if (eventChangeInfo.MinutesHasChanged)
            {
                detectedChange = true;
                result.Add(NotificationsMessages.Meeting.ModificarMinutes, [BodyRole.AllInvitedUsers, BodyRole.Scheduler]);
            }

            //Another Change
            //TODO esto debe ir aqui
            if (!detectedChange && eventChangeInfo.AnyChange)
            {
                result.Add(NotificationsMessages.Meeting.Modificar, [BodyRole.AllInvitedUsers, BodyRole.Scheduler]);
            }
        }

        return result;
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
            {ReplacedTokens.Location, publishedVersion.GetUbicacionCompleta()},
            {"SharedEventId", eventChangeInfo.SharedEventId }
        };
        if (eventChangeInfo.OrderOfDayHasChanged)
        {
            valuesReplace.AddWithCheck(ReplacedTokens.AgendaTitle, eventChangeInfo.ModifiedOrderOfDayName);
        }
        return valuesReplace;
    }

    private async Task UpdateOnlineMeeting(IPnPContext ctx, EventSendOperationInformation eventChangeInfo, MeetingEnEdicionDAO currentState)
    {
        using var graphHelper = new GraphHelper(ctx);

        var eventIsCelebrating = eventChangeInfo.NewEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InCelebration) || eventChangeInfo.CurrentEventStatusId.Equals(TaxonomyValuesIds.EventStatus.InCelebration);

        if (!eventIsCelebrating || !RequiredTeamsOnlineMeetingCreation(currentState))
        {
            Log.LogDebug("Event is not in celebration or does not require online meeting creation");
            return;
        }

        var onlineMeeting = await EnsureEventOnlineMeeting(ctx, graphHelper, currentState);

        var testsenters = await _roleSrv.GetMembersOfBodyByIdAndRole(eventChangeInfo.BodyId, [BodyRole.Scheduler, BodyRole.SchedulerAssistant]);
        var allAttendees = eventChangeInfo.AllAttendees.Concat(testsenters).ToHashSet().ToList();

        var upnToUpnMapping = await MeetingParticipantsHelper.GetUpnToMeetingParticipantUpn(ctx, allAttendees);

        var onlineMeetingParticipants = new List<MeetingParticipantInfo>();

        // for each attendee get all the multi tenant upns and add them as attendees so user could join from any tenant
        foreach (var attendee in allAttendees)
        {
            var istestsenter = testsenters.Any(p => p.Equals(attendee, StringComparison.OrdinalIgnoreCase));
            var multiTenantUpnCollection = upnToUpnMapping[attendee];

            // if user is C or GC add it as testsenter and Coorganizer
            foreach (var multiTenantUpn in multiTenantUpnCollection)
            {
                onlineMeetingParticipants.Add(new MeetingParticipantInfo()
                {
                    Upn = multiTenantUpn,
                    Role = istestsenter ? OnlineMeetingRole.testsenter : OnlineMeetingRole.Attendee
                });
            }
        }

        onlineMeeting.Participants = new MeetingParticipants()
        {
            Attendees = onlineMeetingParticipants
        };

        Log.LogInformation($"Updating online meeting {onlineMeeting.Id} with {onlineMeetingParticipants.Count} participants: {string.Join(", ", onlineMeetingParticipants.Select(p => p.Upn + " (" + p.Role + ")"))}");

        await graphHelper.UpdateOnlineMeeting(_mailBoxUser, onlineMeeting.Id!, onlineMeeting);
    }

    private static bool RequiredTeamsOnlineMeetingCreation(BaseEventDAO CurrentState)
    {
        var eventIsHybridOrOnline = CurrentState.MeetingType?.TermId.ToString().Equals(TaxonomyValuesIds.AttendanceType.Online) == true || CurrentState.MeetingType?.TermId.ToString().Equals(TaxonomyValuesIds.AttendanceType.OnlineInPerson) == true;
        var eventIsTeamsMeeting = CurrentState.OnlineTool?.TermId.ToString().Equals(TaxonomyValuesIds.MeetingTool.Teams) == true;
        var hasExternalJoinUrl = !string.IsNullOrECNTy(CurrentState.UrlOnlineTool?.Url) && string.IsNullOrECNTy(CurrentState.IdReunionOnlineMSTeams);

        return eventIsHybridOrOnline && eventIsTeamsMeeting && !hasExternalJoinUrl;
    }

    private static async Task UpdateAllRelatedDocuments(ILogger log, IPnPContext bodyCtx, Dictionary<Guid, Guid> relatedDocsIdsMappingSrcToDest, Source eventSource, string sharedEventId)
    {
        var eventsproviderPublished = new EventDALprovider(eventSource);
        var agendaItemsprovider = new EventAgendaItemDALprovider(eventsproviderPublished);
        var minutesInfoDAL = new EventMinutesInformationDALprovider(eventsproviderPublished);

        var (_, allPublishedAgendaItems) = await agendaItemsprovider.GetAllDTOByEvent(bodyCtx, sharedEventId);
        foreach (var publishedAgendaItem in allPublishedAgendaItems)
        {
            var relatedDocsIds = new List<string>();
            foreach (var relatedDoc in publishedAgendaItem.RelatedDocumentsIds)
            {
                if (Guid.TryParse(relatedDoc, out var relatedDocGuid)
                    && relatedDocsIdsMappingSrcToDest.TryGetValue(relatedDocGuid, out var dstRelDocId))
                    relatedDocsIds.Add(dstRelDocId.ToString());
                else
                    log.LogWarning($"Related document {relatedDoc} not found in event. Ignoring.");
            }

            if (Guid.TryParse(publishedAgendaItem.AgreementRelatedCertificateId, out var agreementCertId)
                && relatedDocsIdsMappingSrcToDest.TryGetValue(agreementCertId, out var dstCertDocId))
                publishedAgendaItem.AgreementRelatedCertificateId = dstCertDocId.ToString();
            else
                log.LogWarning($"Related certificate document {publishedAgendaItem.AgreementRelatedCertificateId} not found in event. Ignoring.");

            publishedAgendaItem.RelatedDocumentsIds = [.. relatedDocsIds];
        }
        await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, allPublishedAgendaItems);

        var minutesInfoDTO = await minutesInfoDAL.GetBySharedEventId(bodyCtx, sharedEventId);
        if (minutesInfoDTO == null)
            return;
        if (Guid.TryParse(minutesInfoDTO.IdMinutes, out var minutesInfoCertId)
            && relatedDocsIdsMappingSrcToDest.TryGetValue(minutesInfoCertId, out var dstMinutesInfoCertId))
            minutesInfoDTO.IdMinutes = dstMinutesInfoCertId.ToString();
        else
            log.LogWarning($"Related document {minutesInfoDTO.IdMinutes} not found in event. Ignoring.");
        await minutesInfoDAL.AddOrUpdateForEvent(bodyCtx, sharedEventId, minutesInfoDTO);
    }

    [GeneratedRegex(@"(\{){0,1}[0-9A-F]{8}\-[0-9A-F]{4}\-[0-9A-F]{4}\-[0-9A-F]{4}\-[0-9A-F]{12}(\}){0,1}", RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.Singleline, "en-US")]
    private static partial Regex GuidRegEx();
}