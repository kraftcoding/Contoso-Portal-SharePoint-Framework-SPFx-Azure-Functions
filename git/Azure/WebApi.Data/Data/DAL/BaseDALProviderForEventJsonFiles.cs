using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DTO;
using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Data.DAO.Event;

namespace Contoso.Portal.Data.DAL;

public abstract class BaseDALproviderForEventJsonFiles<T, U, V>(EventDALprovider eventsDALprovider) : BaseSPListDALproviderForJsonFiles<T, U, V>
    where T : BaseSPListItemEventDAO, new() // what is stored in the SP list item
    where U : BaseJsonDataDAO, new() // what is stored in the json file
    where V : BaseEventJsonFileDTO<T, U>, new() // what is interchanged with the class consumer
{
    internal readonly EventDALprovider EventDALprovider = eventsDALprovider;

    public async Task<V?> GetDTOByUniqueSharedId(IPnPContext ctx, string uniqueSharedId, string? folderServerRelativeUrl = null)
    {
        var sharedUniqueIdQuery = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery("UniqueSharedID", "Text", uniqueSharedId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, null, folderServerRelativeUrl);
        return (await GetDTOByView(ctx, sharedUniqueIdQuery, folderServerRelativeUrl)).FirstOrDefault();
    }

    public async Task<T?> GetByUniqueSharedId(IPnPContext ctx, string uniqueSharedId, string? folderServerRelativeUrl = null)
    {
        var sharedUniqueIdQuery = PnPContentHelpers.CamlViewBuilder($"<Where><And>{CAML_GetContentTypeFilterPart()}{CAML_GetOperatorQuery("UniqueSharedID", "Text", uniqueSharedId, ComparisonOperators.Eq)}</And></Where>", DefaultViewFields, null, folderServerRelativeUrl);
        return (await GetByView(ctx, sharedUniqueIdQuery, folderServerRelativeUrl)).FirstOrDefault();
    }

    public async Task<(BaseEventDAO Event, IEnumerable<V> Results)> GetAllDTOByEvent(IPnPContext ctx, string sharedEventId)
    {
        var (theEvent, hiddenFolderServerRelativePath) = await EventDALprovider.GetHiddenFolderPathBySharedEventId(ctx, sharedEventId);
        var agendaItems = await GetAllDTOByPath(ctx, hiddenFolderServerRelativePath);
        return (theEvent, agendaItems);
    }

    public async Task<IEnumerable<V>> GetAllDTOByPath(IPnPContext ctx, string folderServerRelativeUrl)
    {
        var view = PnPContentHelpers.CamlViewBuilder($"<Where>{CAML_GetContentTypeFilterPart()}</Where>", DefaultViewFields, DefaultViewSort, folderServerRelativeUrl);
        return await GetDTOByView(ctx, view, folderServerRelativeUrl);
    }

    public async Task<IEnumerable<V>> AddOrUpdateItemsForEvent(IPnPContext ctx, string sharedEventId, IEnumerable<V> items)
    {
        var (_, folder) = await EventDALprovider.EnsureHiddenFolderBySharedEventId(ctx, sharedEventId);
        var results = await AddOrUpdateInFolder(ctx, folder, items);
        return results;
    }

    public async Task<IEnumerable<V>> AddOrUpdateItemsForEvent<C>(IPnPContext ctx, string sharedEventId, IEnumerable<C> items, Func<C, IEnumerable<V>, BaseEventDAO, V?> mapper)
    {
        var (theEvent, resultItems) = await GetAllDTOByEvent(ctx, sharedEventId);
        var selectedItems = new List<V>();
        foreach (var item in items)
        {
            var mappedItem = mapper(item, resultItems, theEvent);
            if (mappedItem != null)
                selectedItems.Add(mappedItem);
        }
        return await AddOrUpdateItemsForEvent(ctx, sharedEventId, selectedItems);
    }

    public async Task<IEnumerable<V>> SelectUpdateByEvent(IPnPContext ctx, string sharedEventId, Func<V, BaseEventDAO, V?> selectChangeDelegate)
    {
        var (theEvent, resultItems) = await GetAllDTOByEvent(ctx, sharedEventId);

        var selectedAndUpdatedItems = resultItems.Select((ai) => { return selectChangeDelegate(ai, theEvent); }).WhereNotNull();

        return await AddOrUpdateItemsForEvent(ctx, sharedEventId, selectedAndUpdatedItems);
    }
}