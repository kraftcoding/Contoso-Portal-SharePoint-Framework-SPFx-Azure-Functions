using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Event;

namespace Contoso.Portal.Data.DAL.Tasks;

public class TasksAttendanceDALprovider(InConstructionEventsDALprovider inConstructionEventsDALprovider) : BaseSPListDALproviderForJsonFiles<TasksAttendanceDAO, TasksAttendanceJsonDAO, TasksAttendanceDTO>
{
    private readonly InConstructionEventsDALprovider _inConstructionEventsDALprovider = inConstructionEventsDALprovider;

    public override string ListSiteRelavteUrl => ListsSiteRelativeUrls.InConstructionEvents;

    public override string DefaultView => PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort);

    public async Task<IEnumerable<TasksAttendanceDTO>> GetAllAtendanceTaskByEvent(IPnPContext ctx, string sharedEventId, bool fromArchive = false)
    {

        var (_, hiddenFolderServerRelativePath) = await (fromArchive ? new ArchivedEventsDALprovider().GetHiddenFolderPathBySharedEventId(ctx, sharedEventId) : _inConstructionEventsDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId));
        var result = await GetAllAtendanceTaskByPath(ctx, hiddenFolderServerRelativePath);
        return result;
    }

    public async Task<TasksAttendanceDTO> GetTaskById(IPnPContext ctx, int taskId)
    {
        return await GetDTOById(ctx, taskId);
    }

    public async Task<IEnumerable<TasksAttendanceDTO>> AddOrUpdateByEvent(IPnPContext ctx, string shareEventId, IEnumerable<TasksAttendanceDTO> taskDTOs)
    {
        var (_, folder) = await _inConstructionEventsDALprovider.EnsureHiddenFolderBySharedEventId(ctx, shareEventId);
        return await AddOrUpdateInFolder(ctx, folder, taskDTOs);
    }

    public async Task<IEnumerable<TasksAttendanceDTO>> GetBySharedEventIdAndAssigned(IPnPContext ctx, string sharedEventId, int userId)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where><And><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksDTO.IdMeeting), "Text", sharedEventId, ComparisonOperators.Eq)}</And>{CAML_GetOperatorQuery(nameof(TasksDTO.AsignadoA), "Integer", userId.ToString(), ComparisonOperators.Eq, true)}</And></Where>", DefaultViewFields, DefaultViewSort);
        return await GetDTOByView(ctx, view);
    }

    public async Task DeleteByEventId(IPnPContext ctx, int eventId, string attendance)
    {
        var (_, hiddenFolderServerRelativePath) = await _inConstructionEventsDALprovider.GetHiddenFolderPathById(ctx, eventId);

        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery(nameof(TasksAttendanceDTO.AsignadoA), "Text", attendance, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, DefaultViewSort, hiddenFolderServerRelativePath);
        var tasks = await GetByView(ctx, view, hiddenFolderServerRelativePath);

        foreach (var att in tasks)
            await att.AsListItem().DeleteAsync();
    }

    private async Task<IEnumerable<TasksAttendanceDTO>> GetAllAtendanceTaskByPath(IPnPContext ctx, string folderServerRelativeUrl)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort, folderServerRelativeUrl);
        return await GetDTOByView(ctx, view, folderServerRelativeUrl);
    }
}