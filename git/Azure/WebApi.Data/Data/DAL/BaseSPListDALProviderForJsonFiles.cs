using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DTO;
using Contoso.Portal.Data.Extensions;
using Newtonsoft.Json;
using PnP.Core;
using PnP.Core.Model;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using System.Text;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAL;

public abstract class BaseSPListDALproviderForJsonFiles<T, U, V> : BaseSPListDALprovider<T>
    where T : BaseSPListItemDAO, new() // what is stored in the SP list item
    where U : BaseJsonDataDAO, new() // what is stored in the json file
    where V : JsonFileListItemAbstractionDTO<T, U>, new() // what is interchanged with the class consumer
{
    public async Task<IEnumerable<V>> GetDTOByView(IPnPContext ctx, string view, string? folderServerRelativeUrl = null)
    {
        // get items from event folder
        var result = await base.GetByView(ctx, view, folderServerRelativeUrl);
        var result2 = result.Where(x => x.FileRef.Contains("6a52_Tes Convoca"));
        // download all file contents and parse the inner json in parallel? 
        var fileResults = result.Select(r => GetAbstractionDTO(r)).ToArray();
        Task.WaitAll(fileResults);

        return fileResults.Select(r => r.Result);
    }

    public async Task<V> GetDTOByUniqueId(IPnPContext ctx, string id, string? folderServerRelativeUrl = null)
    {
        var result = (await base.GetByView(ctx, PnPContentHelpers.CamlViewBuilder($"<Where><Eq><FieldRef Name='UniqueId'/><Value Type='Guid'>{id}</Value></Eq></Where>", DefaultViewFields, null, folderServerRelativeUrl))).First();
        return await GetAbstractionDTO(result);
    }

    public async Task<V> GetDTOById(IPnPContext ctx, int id)
    {
        var result = (await base.GetByView(ctx, PnPContentHelpers.CamlViewBuilder($"<Where><Eq><FieldRef Name='ID'/><Value Type='Counter'>{id}</Value></Eq></Where>", DefaultViewFields))).First();
        return await GetAbstractionDTO(result);
    }


    private async Task<V> GetAbstractionDTO(T result)
    {
        U? jsonDAO = null;
        // if file size is less that 2, then there is no file content to download
        if (int.TryParse((string)result.AsListItem().Values[Fields.FileSize], out var sizeInBytes) && sizeInBytes > 2)
        {
            try
            {
                var fileBytes = await result.AsListItem().File.GetContentBytesAsync();
                jsonDAO = JsonConvert.DeserializeObject<U>(Encoding.UTF8.GetString(fileBytes));
            }
            catch (ServiceException ex)
            {
                throw new Exception($"Error downloading or parsing file content for item '{result.ID}':'{result.AsListItem().File.ServerRelativeUrl}': {ex.Error}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error downloading or parsing file content for item '{result.ID}':'{result.AsListItem().File.ServerRelativeUrl}': {ex.Message}", ex);
            }
        }
        return new V()
        {
            ItemDAO = result,
            JsonDAO = jsonDAO ?? new U()
        };
    }

    public async Task<IEnumerable<V>> AddOrUpdateInLibrary(IPnPContext ctx, IEnumerable<V> entities)
    {
        var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(ListSiteRelavteUrl), p => p.RootFolder.Queryproperties(p => p.ServerRelativeUrl), list => list.ContentTypes.Queryproperties(ct => ct.Id, ct => ct.Fields.Queryproperties(f => f.InternalName, f => f.DefaultValue, f => f.TypeAsString, f => f.FieldTypeKind)));
        var rootFolder = list.RootFolder;
        var contentTypeId = await GetContentTypeIdValueInList(ctx);

        // get items to update
        var itemsToUpdate = entities.Where(ai => ai.Id != 0).ToList();
        if (itemsToUpdate.Any())
            await UpdateItems(ctx, rootFolder.Files, itemsToUpdate, contentTypeId);

        // save the json file and list props
        var itemsToCreate = entities.Where(ai => ai.Id == 0).ToList();
        if (itemsToCreate.Any())
            await AddItems(ctx, rootFolder.Files, itemsToCreate, contentTypeId);

        return entities;
    }

    public async Task<IEnumerable<V>> AddOrUpdateInFolder(IPnPContext ctx, IFolder folder, IEnumerable<V> entities)
    {
        var contentTypeId = await GetContentTypeIdValueInList(ctx);
        var results = new List<V>();

        // get items to update
        var itemsToUpdate = entities.Where(ai => ai.Id != 0).ToList();
        if (itemsToUpdate.Any())
        {
            itemsToUpdate = await RetreiveNeededItemsFromSP(ctx, itemsToUpdate);
            results.AddRange(await UpdateItems(ctx, folder.Files, itemsToUpdate, contentTypeId));
        }

        // save the json file and list props
        var itemsToCreate = entities.Where(ai => ai.Id == 0).ToList();
        if (itemsToCreate.Any())
            results.AddRange(await AddItems(ctx, folder.Files, itemsToCreate, contentTypeId));

        return results;
    }

    public async Task<IEnumerable<V>> AddItems(IPnPContext ctx, IFileCollection eventFiles, List<V> itemsToCreate, string listContentTypeId)
    {
        try
        {
            Task.WaitAll(itemsToCreate.Select(async item =>
            {
                using var mem = new MemoryStream(Encoding.UTF8.GetBytes(item.JsonDAO.ToJsonString()));
                var file = await eventFiles.AddAsync(item.JsonDAO.GetFileName(), mem);

                // save data in list item
                await file.ListItemAllFields.LoadAsync(); // required
                item.ItemDAO.ContentTypeId = listContentTypeId;
                item.ItemDAO.MapToSPListItem(file.ListItemAllFields);
                await file.ListItemAllFields.UpdateOverwriteVersionAsync();
                item.ItemDAO.ID = file.ListItemAllFields.Id;
            }).ToArray());

            // get all new items in batch
            // todo: check if this is needed / optimization
            var savedItemDAOs = await GetByIds(ctx, itemsToCreate.Select(nai => nai.Id).ToArray()); // required to update the _listItem property (otherwise author etc will be null)
            foreach (var item in itemsToCreate)
                item.ItemDAO = savedItemDAOs.First(sai => sai.ID == item.Id);

            return itemsToCreate;
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error adding files: {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding files: {ex.Message}", ex);
        }
    }

    private async Task<List<V>> RetreiveNeededItemsFromSP(IPnPContext ctx, List<V> itemsToCheck)
    {
        try
        {
            var itemsToRetreive = itemsToCheck.Where(item => !item.ItemDAO.HasSPListItem()).ToList();
            if (itemsToRetreive.Any())
            {
                var itemResults = itemsToRetreive.Select(item => GetDTOById(ctx, item.Id)).ToArray();
                Task.WaitAll(itemResults);
                var results = itemResults.Select(r => r.Result).ToList();

                foreach (var item in results)
                {
                    int indice = itemsToCheck.FindIndex(e => e.Id == item.Id);
                    if (indice != -1)
                    {
                        itemsToCheck[indice].ItemDAO.SetSPListItem(item.ItemDAO.GetSPListItem());
                    }
                }
            }

            return itemsToCheck;
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error getting items: {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting items: {ex.Message}", ex);
        }
    }

    public async Task<IEnumerable<V>> UpdateItems(IPnPContext ctx, IFileCollection eventFiles, List<V> itemsToUpdate, string listContentTypeId)
    {
        try
        {
            // update list item data as batch
            var newBatch = new Batch();
            // save the original editor from the context
            Batch? systemUpdateBatch = null;
            FieldUserValue? editorFieldValue = null;
            var userUPN = await ctx.GetCurrentUserUpn(false);
            if (ctx.properties.TryGetValue("IsElevated", out object? isElevated) && (bool)isElevated == true)
            {
                editorFieldValue = new FieldUserValue(await ctx.Web.EnsureUserAsync(userUPN));
                systemUpdateBatch = new Batch();
            }

            var fileUpdates = new List<(string Filename, string Content)>(itemsToUpdate.Count);
            foreach (var item in itemsToUpdate)
            {
                // update list item properties
                item.ItemDAO.ContentTypeId = listContentTypeId;
                var listItem = item.ItemDAO.AsListItem();
                await listItem.UpdateBatchAsync(newBatch);

                // save json file content to update later
                var filePath = listItem.Values[Fields.FileRef]?.ToString() ?? throw new Exception($"Error getting file name for item '{listItem.Id}'");
                var oldJsonIsECNTy = !(int.TryParse((string)listItem.Values[Fields.FileSize], out var sizeInBytes) && sizeInBytes > 2);

                // don't update eCNTy files
                if (oldJsonIsECNTy && item.JsonDAO.IsECNTy())
                    continue;

                var fileName = filePath.Substring(filePath.LastIndexOf('/') + 1);
                fileUpdates.Add((Filename: fileName, Content: item.JsonDAO.ToJsonString()));

                if (systemUpdateBatch != null)
                {
                    listItem[Fields.Editor] = editorFieldValue;
                    await listItem.UpdateBatchAsync(systemUpdateBatch);
                }
            }
            await ctx.ExecuteAsync(newBatch);

            // update json files
            if (fileUpdates?.Count > 0)
            {
                Task.WaitAll(fileUpdates.Select((file) =>
                                {
                                    using var mem = new MemoryStream(Encoding.UTF8.GetBytes(file.Content));
                                    return eventFiles.AddAsync(file.Filename, new MemoryStream(Encoding.UTF8.GetBytes(file.Content)), true);
                                }).ToArray());
            }
            // update editor if exists
            if (systemUpdateBatch != null)
                await ctx.ExecuteAsync(systemUpdateBatch);

            return itemsToUpdate;
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error updating files: {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error updating files: {ex.Message}", ex);
        }
    }
}