using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using PnP.Core.QueryModel;
using static Contoso.Portal.Data.DAL.DALConstants;
using PnP.Core.Model.SharePoint;
using PnP.Core;
using Contoso.Portal.Model.Events;
using PnP.Core.Model;
using System.Text;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace Contoso.Portal.Data.DAL.Event;

public enum Source
{
    InConstruction,
    Published,
    Archived
}

public class EventDALprovider(Source source) : BaseSPListDALprovider<BaseEventDAO>()
{
    private readonly Source _source = source;

    public override string ListSiteRelavteUrl => _source switch
    {
        Source.InConstruction => ListsSiteRelativeUrls.InConstructionEvents,
        Source.Published => ListsSiteRelativeUrls.PublishedEvents,
        Source.Archived => ListsSiteRelativeUrls.ArchivedEvents,
        _ => throw new NotImplementedException()
    };

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public override IList<(string Field, bool Ascending)> DefaultViewSort { get; } = new List<(string Field, bool Ascending)>()
    {
        { (nameof(BaseEventDAO.StartDate), false) }
    };

    public async Task<IEnumerable<BaseEventDAO>> GetByStartDateBetween(IPnPContext ctx, DateTime? startDate, DateTime? endDate)
    {
        var dateFilterPart = CAML_GetDateRangeFilterPart(nameof(BaseEventDAO.StartDate), startDate, endDate);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{dateFilterPart}</And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(ctx, view);
    }

    public async Task<BaseEventDAO?> GetBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(BaseEventDAO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort);
        var result = (await GetByView(ctx, view)).FirstOrDefault();
        return result;
    }

    public async Task<IEnumerable<BaseEventDAO>> GetEventsByStartDateBetweenAndStatusTaxId(PnPContext bodyCtx, DateTime? startDate, DateTime? endDate, string statusTaxonomyId)
    {
        var wssidTermPublished = await bodyCtx.Web.GetWssIdForTermAsync(statusTaxonomyId);
        var estadoMeeting = new List<int>() { wssidTermPublished };

        var dateFilterPart = CAML_GetDateRangeFilterPart(nameof(BaseEventDAO.StartDate), startDate, endDate);
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}<And>{CAML_GetInOperatorQuery(nameof(BaseEventDAO.EstadoMeeting), "Integer", estadoMeeting)}{dateFilterPart}</And></And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetByView(bodyCtx, view);
    }

    public async Task Delete(IPnPContext ctx, int idEvent)
    {
        var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
        var item = await list.Items.GetByIdAsync(idEvent, li => li.All, li => li.File);
        await item.DeleteAsync();
    }

    public async Task<(MeetingEnEdicionDAO Event, IFolder HiddenFolder)> EnsureHiddenFolderBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var theEvent = await GetBySharedEventId(ctx, sharedEventId) ?? throw new Exception($"Event with shared id {sharedEventId} was not found");
        await theEvent.AsListItem().EnsurepropertiesAsync(p => p.Folder.Queryproperties(p => p.ServerRelativeUrl));

        var folder = await theEvent.AsListItem().Folder.EnsureFolderAsync(HiddenFolderName);

        return (theEvent.AsChildCT<MeetingEnEdicionDAO>(), folder);
    }

    public async Task<(BaseEventDAO Event, string HiddenFolderServerRelatevePath)> GetHiddenFolderPathBySharedEventId(IPnPContext ctx, string sharedEventId)
    {
        var theEvent = await GetBySharedEventId(ctx, sharedEventId) ?? throw new Exception($"Event with shared id {sharedEventId} was not found");
        await theEvent.AsListItem().EnsurepropertiesAsync(p => p.Folder.Queryproperties(p => p.ServerRelativeUrl));
        var result = (theEvent, theEvent.AsListItem().Folder.ServerRelativeUrl.UriCombine(HiddenFolderName));
        return result;
    }

    public async Task<(BaseEventDAO Event, string HiddenFolderServerRelatevePath)> GetHiddenFolderPathById(IPnPContext ctx, int eventId)
    {
        var theEvent = await GetById(ctx, eventId) ?? throw new Exception($"Event with id {eventId} was not found");
        await theEvent.AsListItem().EnsurepropertiesAsync(p => p.Folder.Queryproperties(p => p.ServerRelativeUrl));

        return (theEvent, theEvent.AsListItem().Folder.ServerRelativeUrl.UriCombine(HiddenFolderName));
    }
}

public class InConstructionEventsDALprovider : EventDALprovider
{
    public InConstructionEventsDALprovider() : base(Source.InConstruction) { }

    public override IList<string> DefaultViewFields => BaseSPListItemDAO.AllFieldNames<MeetingEnEdicionDAO>();

    #region Managed mails for events

    public async Task<bool> SaveEmailJson(IPnPContext ctx, string sharedEventId, string jsonContent)
    {
        var theEvent = await GetBySharedEventId(ctx, sharedEventId);
        if (theEvent == null)
            return false;

        await theEvent.AsListItem().EnsurepropertiesAsync(
            p => p.Folder.Queryproperties(f => f.ServerRelativeUrl)
        );

        if (theEvent.AsListItem().Folder == null)
            throw new Exception("The event is not a Document Set and has no folder.");

        var hiddenFolderUrl = theEvent.AsListItem().Folder.ServerRelativeUrl.UriCombine(HiddenFolderName);
    
        var parentFolderUrl = theEvent.AsListItem().Folder.ServerRelativeUrl;
        var parentFolder = await ctx.Web.GetFolderByServerRelativeUrlAsync(parentFolderUrl);
        await parentFolder.LoadAsync(p => p.Folders);
        var hiddenFolder = parentFolder.Folders.AsRequested().FirstOrDefault(
            f => f.Name.Equals(HiddenFolderName, StringComparison.OrdinalIgnoreCase)
        );

        if (hiddenFolder == null)
        {
            hiddenFolder = await parentFolder.Folders.AddAsync(HiddenFolderName);
        }


        var folder = ctx.Web.GetFolderByServerRelativeUrl(hiddenFolderUrl);

        var fileUrl = $"{hiddenFolderUrl}/Email.json";
        var file = await ctx.Web.GetFileByServerRelativeUrlOrDefaultAsync(fileUrl);

        JObject existingJson;

        if (file != null)
        {
            var stream = await file.GetContentAsync();
            using var reader = new StreamReader(stream);
            var jsonText = await reader.ReadToEndAsync();

            existingJson = JObject.Parse(jsonText);
        }
        else
        {
            existingJson = new JObject();

            var newBytes = Encoding.UTF8.GetBytes(jsonContent);
            using var unewStream = new MemoryStream(newBytes);

            await folder.Files.AddAsync("Email.json", unewStream, overwrite: true);

            return true;
        }
        
        var incomingJson = JObject.Parse(jsonContent);

        foreach (var prodp in incomingJson.properties())
        {
            if (existingJson[prodp.Name] != null)
            {
                existingJson[prodp.Name] = prodp.Value;
            }
            else
            {
                existingJson.Add(prodp.Name, prodp.Value);
            }
        }

        var updatedBytes = Encoding.UTF8.GetBytes(existingJson.ToString());
        using var updatedStream = new MemoryStream(updatedBytes);

        await folder.Files.AddAsync("Email.json", updatedStream, overwrite: true);

        return true;
    }

    public async Task<string> GetEmailJson(IPnPContext ctx, string sharedEventId)
    {
        var theEvent = await GetBySharedEventId(ctx, sharedEventId);
        if (theEvent == null)
            return null;

        await theEvent.AsListItem().EnsurepropertiesAsync(
            p => p.Folder.Queryproperties(f => f.ServerRelativeUrl)
        );

        if (theEvent.AsListItem().Folder == null)
            throw new Exception("The event is not a Document Set and has no folder.");

        var hiddenFolderUrl = theEvent.AsListItem().Folder.ServerRelativeUrl.UriCombine(HiddenFolderName);

        var file = await ctx.Web.GetFileByServerRelativeUrlOrDefaultAsync(
            $"{hiddenFolderUrl}/Email.json"
        );

        if (file == null)
            return null;

        using var stream = await file.GetContentAsync();
        using var reader = new StreamReader(stream, Encoding.UTF8);

        return await reader.ReadToEndAsync();
    }

    #endregion

    public async Task<MeetingEnEdicionDAO> Create(IPnPContext ctx, MeetingEnEdicionDAO newEvent)
    {
        var folderName = $"{Guid.NewGuid().ToString("N")[..4]}_{newEvent.Title.TruncateString(11)}"; // should be unique enough

        var (item, _, _) = await PnPContentHelpers.CreateDocumentSet(ctx, ListSiteRelavteUrl, folderName, DAOBaseContentTypeId, newEvent.AsNewListItem());

        var storedEvent = await GetById(ctx, item.Id) ?? throw new Exception($"Event with id {item.Id} was not found after creation");
        return storedEvent.AsChildCT<MeetingEnEdicionDAO>();
    }

    // create a specific dal for event changes
    public async Task<bool> HasChangesToPublishOld(IPnPContext ctx, string sharedEventId)
    {
        var eventToCheck = await GetBySharedEventId(ctx, sharedEventId);
        if (eventToCheck == null) // event does not exist
            return false;

        var lastPublish = eventToCheck.LastPublishedDate;

        // event has not been published yet or has been modified after last publish
        if (lastPublish == null || eventToCheck.Modified > lastPublish)
            return true;

        var isAgendaItem = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.AgendaItem, ComparisonOperators.BeginsWith);
        var isDocument = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.DocumentEnConstruccion, ComparisonOperators.BeginsWith);
        var isAgreement = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.Acuerdo, ComparisonOperators.BeginsWith);
        var isMinuteData = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.MinutesInformation, ComparisonOperators.BeginsWith);
        var modifiedAfterLastPublish = CAML_GetDateRangeFilterPart(nameof(BaseSPListItemDAO.Modified), lastPublish, null);

        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
            var result = await list.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
            {
                ViewXml = PnPContentHelpers.CamlViewBuilder($"<Where><And>{modifiedAfterLastPublish}<Or>{isDocument}<Or>{isMinuteData}<Or>{isAgendaItem}{isAgreement}</Or></Or></Or></And></Where>", ["Title", "Id", "ContentTypeId"]),
                FolderServerRelativeUrl = eventToCheck.FileRef
            }, (r) => r);

            return result.Any();
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{sharedEventId}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{sharedEventId}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
        }
    }

    public async Task<(bool, List<PendingChanges>)> HasChangesToPublish(IPnPContext ctx, string sharedEventId, BaseEventDAO eventToCheck, int versionId)
    {
        try
        {
            if (eventToCheck == null)
                return (false, []);
            var pendingChanges = await GetPendingChanges(ctx, eventToCheck, eventToCheck.LastPublishedDate, versionId);
            return (pendingChanges.Count != 0, pendingChanges);
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{sharedEventId}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{sharedEventId}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
        }
    }

    private async Task<List<PendingChanges>> GetPendingChanges(IPnPContext ctx, BaseEventDAO eventToCheck, DateTime? lastPublish, int versionLabel)
    {
        var pendingChanges = new List<PendingChanges>();
        var isAgendaItem = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.AgendaItem, ComparisonOperators.BeginsWith);
        var isDocument = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.DocumentEnConstruccion, ComparisonOperators.BeginsWith);
        var isAgreement = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.Acuerdo, ComparisonOperators.BeginsWith);
        var isMinuteData = CAML_GetOperatorQuery(nameof(BaseSPListItemDAO.ContentTypeId), "ContentTypeId", ContentTypeIds.MinutesInformation, ComparisonOperators.BeginsWith);
        var modifiedAfterLastPublish = CAML_GetDateRangeFilterPart(nameof(BaseSPListItemDAO.Modified), lastPublish, null);

        try
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var (currentVersion, publishedVersion, changedFields) = await eventDAL.GetChangesFromVersion(ctx, eventToCheck.ID, versionLabel);

            var eventChanges = new EventChanges();
            var pendingChangesEvent = new List<PendingChanges>();
            if (changedFields is not null)
            {
                var pendingChangesDALprovider = new EventPendingChangeDALprovider();
                eventChanges = pendingChangesDALprovider.FilterEventChanges(currentVersion, changedFields, publishedVersion);
                pendingChangesEvent = eventChanges is not null ? pendingChangesDALprovider.processEventChanges(eventChanges) : [];
                pendingChanges.AddRange(pendingChangesEvent);

                if (changedFields.Contains(nameof(MeetingEnEdicionDAO.Asistentes)))
                {
                    var toAttendees = currentVersion?.Asistentes?.Select(a => a.AsUserPrincipalName()).WhereNotNull() ?? [];
                    var fromAttendees = publishedVersion?.Asistentes?.Select(a => a.AsUserPrincipalName()).WhereNotNull() ?? [];

                    var attendees = new AttendeesChangesResult()
                    {
                        Added = toAttendees.Except(fromAttendees).ToList(),
                        Removed = fromAttendees.Except(toAttendees).ToList(),
                        NotChanged = toAttendees.Intersect(fromAttendees).ToList()
                    };
                    if (attendees is not null)
                    {
                        var attendanceChanges = FilterAddOrRemovedAttendance(attendees);
                        var pendingChangesAttendance = attendanceChanges is not null ? EventPendingChangeDALprovider.FilterAttendance(attendanceChanges) : [];
                        pendingChanges.AddRange(pendingChangesAttendance);
                    }
                }
            }

            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.Title, p => p.Fields.Queryproperties(p => p.InternalName, p => p.FieldTypeKind, p => p.TypeAsString, p => p.Title));
            var viewXML = PnPContentHelpers.CamlViewBuilder($"<Where><And>{modifiedAfterLastPublish}<Or>{isDocument}<Or>{isMinuteData}<Or>{isAgendaItem}{isAgreement}</Or></Or></Or></And></Where>", ["Title", "Id", "ContentTypeId", "ContentType", "FileLeafRef"]);
            var result = await list.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
            {
                ViewXml = viewXML,
                FolderServerRelativeUrl = eventToCheck.FileRef
            }, (r) => r);

            if (result == null || !result.Any())
                return pendingChanges;
            //Filter changes for documents
            string idValues = string.Join("", result.Where(item => item.ContentType.Id.StartsWith(ContentTypeIds.DocumentEnConstruccion)).Select(x => $"<Value Type='Number'>{x.Id}</Value>"));
            var documentsViewXml = PnPContentHelpers.CamlViewBuilder($"<Where><And><BeginsWith><FieldRef Name='ContentTypeId'/><Value Type='ContentTypeId'>{ContentTypeIds.DocumentEnConstruccion}</Value></BeginsWith><In><FieldRef Name='ID'/><Values>{idValues}</Values></In></And></Where>", ["Title", "Id", "ContentTypeId", "ContentType", "FileLeafRef", "Editor", nameof(IEventDocument.SensitivityLabelprocessing), nameof(IEventDocument.SensitivityLabelIsSystem)]);
            var documentsResult = await list.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
            {
                ViewXml = documentsViewXml,
                FolderServerRelativeUrl = eventToCheck.FileRef
            }, (r) => r);
            var documentsToFilter = documentsResult.Where(d => ((bool)d[nameof(IEventDocument.SensitivityLabelprocessing)]) == false && (((FieldUserValue)d["Editor"]).LookupValue != "Aplicación de SharePoint" || ((bool)d[nameof(IEventDocument.SensitivityLabelIsSystem)]) == false));
            var documents = FilterPendingChanges(documentsToFilter, "Document") ?? [];
            documents = documents.Where(d => d.ChangeValue != "Email.json").ToList(); //AMC: Excluimos Email.json, para que no identifique que ha habido un cambio en alguno de los Documents de la Meeting por el mero hecho de que se haya modificado el JSON de gestión de emails, ya que este cambio no afecta a la Meeting en sí misma

            //Filter changes for agenda items
            var agendaItems = FilterPendingChanges(result.Where(item => item.ContentType.Id.StartsWith(ContentTypeIds.AgendaItem)), "AgendaItem") ?? [];

            //Filter changes for agreements
            var agreements = FilterPendingChanges(result.Where(item => item.ContentType.Id.StartsWith(ContentTypeIds.Acuerdo)), "Agreement") ?? [];

            pendingChanges.AddRange(documents);
            pendingChanges.AddRange(agendaItems);
            pendingChanges.AddRange(agreements);
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{eventToCheck.ID}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting event changes '{ListSiteRelavteUrl}' '{eventToCheck.ID}' in site '{ctx.Uri.AbsolutePath}': {ex.Message}", ex);
        }
        return pendingChanges;
    }

    private static List<AttendanceChanges> FilterAddOrRemovedAttendance(AttendeesChangesResult attendees)
    {
        var attendanceChanges = new List<AttendanceChanges>();

        if (attendees != null)
        {
            attendanceChanges.AddRange(attendees.Added.Select(user => new AttendanceChanges { AddOrRemoved = "AddAttendance", user = user }));
            attendanceChanges.AddRange(attendees.Removed.Select(user => new AttendanceChanges { AddOrRemoved = "RemoveAttendance", user = user }));
        }

        return attendanceChanges;
    }

    private static List<PendingChanges> FilterPendingChanges(IEnumerable<IListItem> result, string type)
    {
        var pendingChanges = new List<PendingChanges>();
        if (result.Any())
        {
            foreach (var item in result)
            {
                var title = type == "Document" ? item.Values[Fields.FileLeafRef] as string : item.Title;
                pendingChanges.Add(new PendingChanges
                {
                    ChangeType = type,
                    ChangeValue = title ?? string.ECNTy
                });
            }
        }
        return pendingChanges;
    }
}

public class PublishedEventsDALprovider : EventDALprovider
{
    public PublishedEventsDALprovider() : base(Source.Published) { }

    public override IList<string> DefaultViewFields => BaseSPListItemDAO.AllFieldNames<MeetingPublicadaDAO>();

    public async Task<MeetingPublicadaDAO> CreateOrUpdateFrom(IPnPContext ctx, string fileLeafRef, MeetingEnEdicionDAO publishedVersion)
    {
        var publishedproperties = new MeetingPublicadaDAO()
        {
            Title = publishedVersion.Title,
            Location = publishedVersion.Location,
            LocationDetails = publishedVersion.LocationDetails,
            DocumentSetDescription = publishedVersion.DocumentSetDescription,
            EndDate = publishedVersion.EndDate,
            StartDate = publishedVersion.StartDate,
            OutlookMeetingId = publishedVersion.OutlookMeetingId,
            IdReunionOnlineMSTeams = publishedVersion.IdReunionOnlineMSTeams,
            IdMeeting = publishedVersion.IdMeeting,
            MeetingType = publishedVersion.MeetingType.Clone(),
            OnlineTool = publishedVersion.OnlineTool.Clone(),
            EstadoMeeting = publishedVersion.EstadoMeeting.Clone(),
            UrlOnlineTool = publishedVersion.UrlOnlineTool.Clone(),
            Asistentes = publishedVersion.Asistentes.WhereNotNull().Select(a => a.Clone()!).ToArray(),
            GrupoAsistentesAsociado = publishedVersion.GrupoAsistentesAsociado.Clone(),
            VersionPublicada = publishedVersion.VersionPublicada,
            LastPublishedDate = publishedVersion.LastPublishedDate
        };

        var (item, _, _) = await PnPContentHelpers.CreateDocumentSet(ctx, ListSiteRelavteUrl, fileLeafRef, await GetContentTypeIdValueInList(ctx), publishedproperties.AsNewListItem(), true);
        var storedEvent = await GetById(ctx, item.Id) ?? throw new Exception($"Event with id {item.Id} was not found after creation");
        return storedEvent.AsChildCT<MeetingPublicadaDAO>();
    }
}

public class ArchivedEventsDALprovider : EventDALprovider
{
    public ArchivedEventsDALprovider() : base(Source.Archived) { }

    public override IList<string> DefaultViewFields => BaseSPListItemDAO.AllFieldNames<MeetingPublicadaDAO>();
}