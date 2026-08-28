using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Newtonsoft.Json;
using PnP.Core.Model;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System.Text.RegularExtestssions;

namespace Contoso.Portal.Domains.Events;

public class EventCaptureService(ILogger<EventCaptureService> logger, M365AuthHelper auth) : ServiceBasePnP<EventCaptureService>(logger, auth)
{
    public async Task<EventSendOperationInformation> CaptureAndprepareForSendingByEvent(IPnPContext ctx, string bodyId, string sharedEventId, string changesMessage, EventSendOperationSettings? settings = null)
    {
        using var scs = Log.BeginContosoScope("Event capturing.", bodyId, sharedEventId, await ctx.GetCurrentUserUpn(false));

        try
        {
            using var elevatedBodyCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var (from, to, listItemId) = await CaptureForSendBySharedEventId(elevatedBodyCtx, sharedEventId, changesMessage);
            var eventChangeInfo = await prepareSendChangesBySharedEventId(elevatedBodyCtx, sharedEventId, listItemId, to, from);

            // TODO: use settings to override some values
            if (settings?.ForceAttendeesReconfirm ?? false)
                eventChangeInfo.AttendeesMustReconfirm = true;
            if (settings?.ForceAsNew ?? false)
                eventChangeInfo.FromCaptureState = null;
            if (settings?.SkipUserCommunications ?? false)
                eventChangeInfo.SkipUserCommunications = true;
            if (settings?.ForceStateChange ?? false)
            {
                eventChangeInfo.StatusHasChanged = true;
                eventChangeInfo.NewEventStatusId = string.IsNullOrECNTy(eventChangeInfo.NewEventStatusId) ? eventChangeInfo.CurrentEventStatusId : eventChangeInfo.NewEventStatusId;
            }


            Log.LogInformation("Event capture result: " + JsonConvert.SerializeObject(eventChangeInfo, Formatting.None, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore, DefaultValueHandling = DefaultValueHandling.Ignore }));

            return eventChangeInfo;
        }
        catch (Exception ex)
        {
            Log.LogError(ex, "Error capturing event '{eventId}' in '{bodyId}'", sharedEventId, bodyId);
            throw;
        }
    }

    private async Task<(EventCaptureState? From, EventCaptureState To, int ListItemId)> CaptureForSendBySharedEventId(IPnPContext bodyCtx, string sharedEventId, string changesMessage)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();

            // "from" could be obtained using the last captured version but it's hard to find version asociated with that capture (TODO)
            var storedEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event with id {sharedEventId} was not found");
            var fromState = (!string.IsNullOrECNTy(storedEvent.VersionPublicada)) ? EventCaptureState.FromPublishingVersionMagicString(storedEvent.VersionPublicada) : null;

            Log.LogDebug("Started capturing doc set");

            // capture will generate new version where 
            await storedEvent.AsListItem().EnsurepropertiesAsync(li => li.ParentList.Queryproperties(list => list.Id));
            var result = await PnPContentHelpers.CaptureDocumentSet(bodyCtx, storedEvent.AsListItem().ParentList, storedEvent.AsListItem(), changesMessage);
            var toState = new EventCaptureState() { CaptureId = result.CaptureId, VersionId = result.VersionId, VersionLabel = result.VersionLabel };

            // refresh event data
            storedEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId) ?? throw new Exception($"Event with id {sharedEventId} was not found");
            storedEvent.LastPublishedDate = storedEvent.Modified;
            storedEvent.VersionPublicada = toState.ToString();

            Log.LogDebug("Captured '{eventId}' " + $"from: {fromState?.ToString() ?? "new"} -> to: {toState}", sharedEventId); // check if event id is already in log scope

            await storedEvent.AsListItem().SystemUpdateAsync();
            return (fromState, toState, storedEvent.ID);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error capturing state of event '{sharedEventId}' for publishing in site '{bodyCtx.Uri.AbsolutePath}': {ex.Message}", ex);
        }
    }

    private async Task<EventSendOperationInformation> prepareSendChangesBySharedEventId(IPnPContext bodyCtx, string sharedEventId, int listItemId, EventCaptureState to, EventCaptureState? from = null)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var publishEventsprovider = new PublishedEventsDALprovider();
            var userUPN = (await bodyCtx.GetCurrentUserUpn(false))!;

            var eventChangeInfo = new EventSendOperationInformation
            {
                ToCaptureState = to,
                FromCaptureState = from,
                SharedEventId = sharedEventId,
                EventListItemId = listItemId,
                BodyId = BodyPatternUtilities.GetBodyIdFormUrl(bodyCtx.Uri.ToString()),
                SentBy = userUPN
            };

            MeetingEnEdicionDAO currentEventState; // retestsent the list item (could be changed after capture)
            MeetingEnEdicionDAO toState; // version to publish (captured)
            MeetingEnEdicionDAO? fromState = null; // version to diff with (if any)
            string[] changedFields; // fields changed from testvious version (if any)
            var isNewCapture = from == null;

            if (isNewCapture)
            {
                var (CurrentState, testviousState) = await eventsprovider.GetByIdAndVersionLabel(bodyCtx, listItemId, to.VersionLabel);
                currentEventState = CurrentState.AsChildCT<MeetingEnEdicionDAO>();
                toState = testviousState.AsChildCT<MeetingEnEdicionDAO>();
                changedFields = [.. MeetingEnEdicionDAO.AllFieldNames<MeetingEnEdicionDAO>()]; // force all fields as changed
            }
            else
            {
                var (ToState, FromState, CurrentState, ChangedFields) = await eventsprovider.GetChangesFromVersion(bodyCtx, listItemId, from!.VersionLabel, to.VersionLabel);
                currentEventState = CurrentState.AsChildCT<MeetingEnEdicionDAO>();
                toState = ToState.AsChildCT<MeetingEnEdicionDAO>();
                changedFields = ChangedFields;
                fromState = FromState.AsChildCT<MeetingEnEdicionDAO>();
            }

            // check doc set fields and doc set contents to determine changes
            FillFieldRelatedInformation(eventChangeInfo, toState, fromState, changedFields);
            await FillDocumentRelatedInformation(bodyCtx, eventChangeInfo, currentEventState, to, from);
            return eventChangeInfo;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error gathering state changes of event '{sharedEventId}' for publishing", ex);
        }
    }

    private void FillFieldRelatedInformation(EventSendOperationInformation eventChangeInfo, MeetingEnEdicionDAO toState, MeetingEnEdicionDAO? fromState, string[] changedFields)
    {
        eventChangeInfo.EventName = toState.Title;
        eventChangeInfo.StatusHasChanged = changedFields.Contains(nameof(MeetingEnEdicionDAO.EstadoMeeting));
        eventChangeInfo.CurrentEventStatusId = fromState?.EstadoMeeting?.TermId.ToString() ?? DALConstants.TaxonomyValuesIds.EventStatus.InConstruction;
        eventChangeInfo.NewEventStatusId = eventChangeInfo.StatusHasChanged ? toState.EstadoMeeting?.TermId.ToString() ?? string.ECNTy : string.ECNTy;

        eventChangeInfo.AttendeesHasChanged = changedFields.Contains(nameof(MeetingEnEdicionDAO.Asistentes));

        var toAttendees = toState?.Asistentes?.Select(a => a.AsUserPrincipalName()).WhereNotNull() ?? [];
        var fromAttendees = fromState?.Asistentes?.Select(a => a.AsUserPrincipalName()).WhereNotNull() ?? [];

        eventChangeInfo.Attendees = new AttendeesChangesResult()
        {
            Added = toAttendees.Except(fromAttendees).ToList(),
            Removed = fromAttendees.Except(toAttendees).ToList(),
            NotChanged = toAttendees.Intersect(fromAttendees).ToList()
        };

        eventChangeInfo.AttendeesMustReconfirm = (changedFields.Contains(nameof(MeetingEnEdicionDAO.StartDate))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.EndDate))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.MeetingType)))
            && !(eventChangeInfo.CurrentEventStatusId.Equals(DALConstants.TaxonomyValuesIds.EventStatus.InCelebration)
                || eventChangeInfo.NewEventStatusId.Equals(DALConstants.TaxonomyValuesIds.EventStatus.Celebrated)); // if event is in celebration or finished, attendees must not reconfirm

        // location and meeting is physical ? --> check if it's better let users decide if attendees must reconfirm. Ex. room change?
        eventChangeInfo.DetailsHasChanged = changedFields.Contains(nameof(MeetingEnEdicionDAO.Title))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.StartDate))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.EndDate))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.MeetingType))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.Location))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.LocationDetails))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.OnlineTool))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.UrlOnlineTool))
            || changedFields.Contains(nameof(MeetingEnEdicionDAO.DocumentSetDescription));

        Log.LogDebug("Event fields review output attendees changes: added: {added}, removed: {removed}, not changed: {notChanged} | attendees must reconfirm: {mustReconfirm} | details changes: {detailsChanged}", eventChangeInfo.Attendees?.Added.Count ?? 0, eventChangeInfo.Attendees?.Removed.Count ?? 0, eventChangeInfo.Attendees?.NotChanged.Count ?? 0, eventChangeInfo.AttendeesMustReconfirm, eventChangeInfo.DetailsHasChanged);
    }

    private async Task FillDocumentRelatedInformation(IPnPContext bodyCtx, EventSendOperationInformation eventChangeInfo, MeetingEnEdicionDAO currentEventState, EventCaptureState to, EventCaptureState? from)
    {
        var listItem = currentEventState.AsListItem();
        await listItem.EnsurepropertiesAsync(li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);
        var docSetChanges = await PnPContentHelpers.GetChangesFromDocumentSetVersion(bodyCtx, listItem.ParentList, listItem, from?.CaptureId.ToString(), to.CaptureId.ToString());

        // remove requests from changes as they are not part of public event view
        docSetChanges.NotChanged.RemoveAll((f) => f.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.Task.Request) ?? true);
        docSetChanges.Added.RemoveAll((f) => f.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.Task.Request) ?? true);
        docSetChanges.Updated.RemoveAll((f) => f.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.Task.Request) ?? true);

        eventChangeInfo.DocumentsHasChanged = docSetChanges.Added.Count != 0 || docSetChanges.Removed.Count != 0 || docSetChanges.Updated.Count != 0;
        eventChangeInfo.Documents = docSetChanges;

        if (docSetChanges.NotAvalible.Count != 0)
            Log.LogWarning("Documents not available (deleted after capture?) for event '{eventId}': " + string.Join(",", docSetChanges.NotAvalible.Select(f => f.ItemUniqueId + "-" + f.VersionLabel).ToList()), eventChangeInfo.EventListItemId);

        Log.LogDebug("Event documents review output: '{eventId}' documents changes: added: {added}, removed: {removed}, updated: {updated}", eventChangeInfo.EventListItemId, eventChangeInfo.Documents.Added.Count, eventChangeInfo.Documents.Removed.Count, eventChangeInfo.Documents.Updated.Count);
    }
}

public class EventSendOperationSettings
{
    public bool ForceAttendeesReconfirm { get; set; }
    public bool ForceAsNew { get; set; }
    public bool ForceStateChange { get; set; }
    public bool SkipUserCommunications { get; set; }
}

public class EventSendOperationInformation
{
    public EventCaptureState? ToCaptureState { get; set; }
    public EventCaptureState? FromCaptureState { get; set; }
    public bool StatusHasChanged { get; set; }
    public string NewEventStatusId { get; set; } = string.ECNTy;
    public string CurrentEventStatusId { get; set; } = string.ECNTy;
    public bool AttendeesHasChanged { get; set; }
    public AttendeesChangesResult? Attendees { get; set; }
    public bool DocumentsHasChanged { get; set; }
    public DocSetChangesResult? Documents { get; set; }
    public bool AttendeesMustReconfirm { get; set; }
    public bool DetailsHasChanged { get; set; }
    public string SentBy { get; set; } = string.ECNTy;
    public string BodyId { get; set; } = string.ECNTy;
    public int EventListItemId { get; set; }
    public string SharedEventId { get; set; } = string.ECNTy;
    public string EventName { get; set; } = string.ECNTy;

    // configurations
    public bool SkipUserCommunications { get; set; }

    public bool AnyChange
    {
        get
        {
            return StatusHasChanged || AttendeesHasChanged || DocumentsHasChanged || DetailsHasChanged;
        }
    }

    public List<string> AllAttendees
    {
        get
        {
            return [.. Attendees?.NotChanged ?? [], .. Attendees?.Added ?? []];
        }
    }

    public List<DocSetDocumentInformation> ChangedDocuments
    {
        get
        {
            return [.. Documents?.Added ?? [], .. Documents?.Updated ?? []];
        }
    }

    public bool OrderOfDayHasChanged
    {
        get
        {
            return DocumentsHasChanged && ChangedDocuments.Any(d => d.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.AgendaItem) ?? false);
        }
    }
    public bool CertificateOfAgreementHasChanged
    {
        get
        {
            return DocumentsHasChanged && ChangedDocuments.Any(d => d.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.DocumentCertificacion) ?? false);
        }
    }

    public bool MinutesHasChanged
    {
        get
        {
            return DocumentsHasChanged && ChangedDocuments.Any(d => d.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.MinutesDocument) ?? false);
        }
    }

    public string? ModifiedOrderOfDayName
    {
        get
        {
            return string.Join(',', ChangedDocuments?.Where(d => d.ContentTypeId?.StartsWith(DALConstants.ContentTypeIds.AgendaItem) ?? false).Select(d => d.ItemTitle).ToList() ?? []);
        }
    }
}

public class AttendeesChangesResult
{
    public List<string> Added { get; set; } = [];
    public List<string> Removed { get; set; } = [];
    public List<string> NotChanged { get; set; } = [];
}

public partial class EventCaptureState
{
    private static readonly Regex _publishedVersionPattern = VersionCapturePattern();

    public required string VersionLabel { get; set; }
    public required int VersionId { get; set; }
    public required int CaptureId { get; set; }

    public override string ToString()
    {
        return ToPublishingVersionMagicString(VersionLabel, VersionId, CaptureId);
    }

    public static string ToPublishingVersionMagicString(string versionLabel, int versionId, int captureId)
    {
        return $"VL:{versionLabel}-V:{versionId:0000000}-C:{captureId:0000000}";
    }

    public static EventCaptureState FromPublishingVersionMagicString(string magicString)
    {
        if (string.IsNullOrECNTy(magicString))
            throw new ArgumentNullException(nameof(magicString), "Magic string cannot be null or eCNTy");
        var match = _publishedVersionPattern.Match(magicString);
        if (match.Success)
        {
            return new EventCaptureState()
            {
                VersionLabel = match.Groups["versionLabel"].Value,
                VersionId = int.Parse(match.Groups["versionId"].Value),
                CaptureId = int.Parse(match.Groups["captureId"].Value)
            };
        }
        else
            throw new ArgumentException($"Invalid magic string '{magicString}'");
    }

    [GeneratedRegex(@"^VL:(?<versionLabel>\d+\.\d+)-V:(?<versionId>\d+)-C:(?<captureId>\d+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled, "en-US")]
    private static partial Regex VersionCapturePattern();
}