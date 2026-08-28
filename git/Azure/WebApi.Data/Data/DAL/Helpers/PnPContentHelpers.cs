using System.Reflection;
using System.Text;
using Contoso.Portal.Data.Extensions;
using PnP.Core;
using PnP.Core.Model;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Helpers;

public static class PnPContentHelpers
{
    public static async Task<IEnumerable<FileInFolderInformation>> GetFolderFilesInformation(IListItem folderItem)
    {
        await folderItem.EnsurepropertiesAsync(fi => fi.Id, fi => fi.FileSystemObjectType, fi => fi.Folder.Queryproperties(f => f.ServerRelativeUrl), fi => fi.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);

        if (folderItem.FileSystemObjectType != FileSystemObjectType.Folder)
            throw new Exception("Item is not a folder");

        return await folderItem.ParentList.GetAllQueryResultsAsEntityAsync(new RenderListDataOptions()
        {
            ViewXml = CamlViewBuilder("<Where><Eq><FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value></Eq></Where>", ["ID", "Title", "UniqueId", "FileRef", "FileLeafRef", "ContentTypeId", "_UIVersionString", "File_x0020_Size"], null, folderItem.Folder.ServerRelativeUrl, ViewScope.Recursive),
            FolderServerRelativeUrl = folderItem.Folder.ServerRelativeUrl
        }, (item) =>
        {
            var filePath = (item.Values["FileRef"] as string)!;
            var filePathDocSetRelative = filePath.Substring(folderItem.Folder.ServerRelativeUrl.Length);

            //coger de aqui para traer la info del doc
            return new FileInFolderInformation()
            {
                ItemTitle = item.Title,
                ItemId = item.Id,
                ItemUniqueId = Guid.Parse((item.Values["UniqueId"] as string)!),
                FileServerRelativePath = filePath,
                FileName = (item.Values["FileLeafRef"] as string)!,
                CurrentVersionLabel = (item.Values["_UIVersionString"] as string)!,
                ContentTypeId = (item.Values["ContentTypeId"] as string)!,
                FileSize = (item.Values["File_x0020_Size"] as string)!,
                FolderRelativeUrl = filePathDocSetRelative,
                BaseListItem = item
            };
        });
    }

    public static async Task<FolderCopyMoveResult> CopyOrMoveFolderContents(IPnPContext ctx, IListItem srcFolder, IListItem dstFolder, bool isMove, MoveCopyOptions? options = null, Func<FileInFolderInformation, bool>? shouldSkipDelegate = null, bool overwrite = true, MoveOperations? moveOperations = null)
    {
        var sourceUniqueIdToDestinationPaths = new Dictionary<Guid, string>();
        var filesSkipped = new List<FileInFolderInformation>();

        await srcFolder.EnsurepropertiesAsync(li => li.FileSystemObjectType, li => li.Folder.Queryproperties(f => f.Name, f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);
        await dstFolder.EnsurepropertiesAsync(li => li.FileSystemObjectType, li => li.Folder.Queryproperties(f => f.Name, f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);

        if (srcFolder.FileSystemObjectType != FileSystemObjectType.Folder || dstFolder.FileSystemObjectType != FileSystemObjectType.Folder)
            throw new Exception("Source or destination is not a folder");

        // get all files from source folder
        var allSrcFiles = await GetFolderFilesInformation(srcFolder);

        // prepare folder structure in destination
        await EnsureFolderStructureForFiles(dstFolder.Folder, allSrcFiles.Select(f => f.FolderRelativeUrl));

        // copy and maintain unique ids to destination path
        var copyBatch = ctx.NewBatch();
        foreach (var file in allSrcFiles)
        {
            if (shouldSkipDelegate is not null && shouldSkipDelegate(file))
            {
                filesSkipped.Add(file);
                continue;
            }

            var oldServerRelativePath = file.FileServerRelativePath;
            var newServerRelativePath = dstFolder.Folder.ServerRelativeUrl.UriCombine(file.FolderRelativeUrl);

            sourceUniqueIdToDestinationPaths.Add(file.ItemUniqueId, file.FolderRelativeUrl);

            var spfile = await ctx.Web.GetFileByServerRelativeUrlAsync(oldServerRelativePath);

            // calculate copy or move options
            var coptions = options ?? new MoveCopyOptions() { ResetAuthorAndCreatedOnCopy = true, ShouldBypassSharedLocks = true };
            var cmoveOperations = (moveOperations is not null) ? moveOperations.Value : MoveOperations.BypassSharedLock;
            cmoveOperations = overwrite ? cmoveOperations & MoveOperations.Overwrite : cmoveOperations;

            if (isMove)
                await spfile.MoveToBatchAsync(copyBatch, newServerRelativePath, cmoveOperations, coptions);
            else
                await spfile.CopyToBatchAsync(copyBatch, newServerRelativePath, overwrite, coptions);
        }
        await ctx.ExecuteAsync(copyBatch);

        // map destinations to unique ids
        var allDstFiles = await GetFolderFilesInformation(dstFolder);
        var relatedDocsIdsMappingSrcToDest = new Dictionary<Guid, Guid>();
        foreach (var dstListItem in allDstFiles)
        {
            var matchingPair = sourceUniqueIdToDestinationPaths.FirstOrDefault((pair) => dstListItem.FolderRelativeUrl.Equals(pair.Value, StringComparison.OrdinalIgnoreCase));
            if (!default(KeyValuePair<Guid, string>).Equals(matchingPair))
                relatedDocsIdsMappingSrcToDest.Add(matchingPair.Key, dstListItem.ItemUniqueId);
        }

        return new FolderCopyMoveResult()
        {
            CopiedSrcToDstIdMapping = relatedDocsIdsMappingSrcToDest,
            AllSrcFiles = [.. allSrcFiles],
            AllDstFiles = [.. allDstFiles],
            SkippedFiles = [.. filesSkipped]
        };
    }

    public static async Task<FolderCopyMoveResult> CopyOrMoveFolder(IPnPContext ctx, IListItem srcFolder, string dstServerRelativeFolderPath, CopyMigrationOptions options)
    {
        var sourceUniqueIdToDestinationPaths = new Dictionary<Guid, string>();
        var filesSkipped = new List<FileInFolderInformation>();

        await srcFolder.EnsurepropertiesAsync(li => li.FileSystemObjectType, li => li.Folder.Queryproperties(f => f.Name, f => f.ServerRelativeUrl), li => li.ParentList.Queryproperties(list => list.Id), li => li.UniqueId);

        if (srcFolder.FileSystemObjectType != FileSystemObjectType.Folder)
            throw new Exception("Source is not a folder");

        // get all files from source folder
        var allSrcFiles = await GetFolderFilesInformation(srcFolder);
        foreach (var file in allSrcFiles)
            sourceUniqueIdToDestinationPaths.Add(file.ItemUniqueId, file.FolderRelativeUrl);

        string destinationAbsoluteUrl = ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(dstServerRelativeFolderPath);
        await ctx.Site.EnsureCopyJobHasFinishedAsync(await ctx.Site.CreateCopyJobsAsync(
            [.. new List<string> { ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(srcFolder.Folder.ServerRelativeUrl) }]
            , destinationAbsoluteUrl, options));

        var dstFolder = await ctx.Web.GetFolderByServerRelativeUrlAsync(dstServerRelativeFolderPath.UriCombine(srcFolder.Folder.Name), (f) => f.ListItemAllFields.Queryproperties(lif => lif.All));

        // map destinations to unique ids
        var allDstFiles = await GetFolderFilesInformation(dstFolder.ListItemAllFields);
        var relatedDocsIdsMappingSrcToDest = new Dictionary<Guid, Guid>();
        foreach (var dstListItem in allDstFiles)
        {
            var matchingPair = sourceUniqueIdToDestinationPaths.FirstOrDefault((pair) => dstListItem.FolderRelativeUrl.Equals(pair.Value, StringComparison.OrdinalIgnoreCase));
            if (!default(KeyValuePair<Guid, string>).Equals(matchingPair))
                relatedDocsIdsMappingSrcToDest.Add(matchingPair.Key, dstListItem.ItemUniqueId);
        }

        return new FolderCopyMoveResult()
        {
            CopiedSrcToDstIdMapping = relatedDocsIdsMappingSrcToDest,
            AllSrcFiles = [.. allSrcFiles],
            AllDstFiles = [.. allDstFiles],
            SkippedFiles = [.. filesSkipped]
        };
    }

    public static async Task<(IListItem item, IFolder folder, IList list)> CreateDocumentSet(IPnPContext ctx, string listSiteRelativeUrl, string documentSetName, string contentTypeId, IDictionary<string, object>? metadata = null, bool shouldOverwrite = false)
    {
        try
        {
            var list = await ctx.Web.Lists.GetByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(listSiteRelativeUrl), p => p.RootFolder.Queryproperties(p => p.ServerRelativeUrl), list => list.ContentTypes.Queryproperties(ct => ct.Id, ct => ct.Fields.Queryproperties(f => f.InternalName, f => f.DefaultValue, f => f.TypeAsString, f => f.FieldTypeKind)));
            var rootFolder = list.RootFolder;
            var ct = list.ContentTypes.AsRequested().First(ct => ct.Id.StartsWith(contentTypeId));

            var folderName = documentSetName.GetSafeName();

            // check if folder already exists
            IFolder folder = await rootFolder.Folders.FirstOrDefaultAsync(f => f.Name == folderName);
            if (folder is null || !folder.Exists)
                folder = await rootFolder.Folders.AddAsync(folderName);
            else if (!shouldOverwrite)
                throw new Exception($"Document set '{documentSetName}' already exists in list '{listSiteRelativeUrl}' in site '{ctx.Uri.AbsolutePath}'");

            folder = await ctx.Web.GetFolderByServerRelativeUrlAsync($"{rootFolder.ServerRelativeUrl}/{folderName}", p => p.ListItemAllFields.Queryproperties(li => li.All,
                        li => li.ParentList.Queryproperties(list => list.Title,
                            list => list.Fields.Queryproperties(f => f.InternalName, f => f.DefaultValue, f => f.TypeAsString, f => f.FieldTypeKind, f => f.Title))));

            folder.ListItemAllFields[DALConstants.Fields.ContentTypeId] = ct.Id;
            // set default values for taxonomy fields (workaround-prodbably not required if it default value prodvisioning issue)
            foreach (var field in ct.Fields.AsRequested().Where(f => !string.IsNullOrECNTy(f.DefaultValue as string) && f.TypeAsString == "TaxonomyFieldType"))
            {
                var termId = (field.DefaultValue as string)?.Split('|').LastOrDefault();
                if (termId is not null)
                    folder.ListItemAllFields[field.InternalName] = new FieldTaxonomyValue(new Guid(termId), string.ECNTy);
            }

            // set addional metadata
            if (metadata is not null)
                foreach (var item in metadata)
                    folder.ListItemAllFields[item.Key] = item.Value;

            await folder.ListItemAllFields.UpdateOverwriteVersionAsync();

            return (folder.ListItemAllFields, folder, list);
        }
        catch (ServiceException ex)
        {
            throw new Exception($"Error creating document set '{documentSetName}' in list '{listSiteRelativeUrl}' in site '{ctx.Uri.AbsolutePath}': {ex.Error}", ex);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error creating document set '{documentSetName}' in list '{listSiteRelativeUrl}' in site '{ctx.Uri.AbsolutePath}'", ex);
        }
    }

    public static async Task<(int VersionId, string VersionLabel, DateTime VersionCreated, int CaptureId)> CaptureDocumentSet(IPnPContext ctx, IList list, IListItem docSetItem, string message, bool shouldCaptureMinor = false)
    {
        using var graphHelper = new GraphHelper(ctx);

        await ctx.Site.EnsurepropertiesAsync(s => s.Id);
        await list.EnsurepropertiesAsync(l => l.Id);

        var captureId = await graphHelper.CaptureDocumentSet(ctx.Site.Id.ToString(), list.Id.ToString(), docSetItem.Id.ToString(), shouldCaptureMinor, message);
        var lastVersion = await docSetItem.Versions.FirstOrDefaultAsync();

        return (lastVersion.Id, lastVersion.VersionLabel, lastVersion.Created, captureId);
    }

    public static async Task<DocSetChangesResult> GetChangesFromDocumentSetVersion(IPnPContext ctx, IList list, IListItem docSetItem, string? fromCaptureId, string toCaptureId)
    {
        using var graphHelper = new GraphHelper(ctx);

        await ctx.Site.EnsurepropertiesAsync(s => s.Id);
        await list.EnsurepropertiesAsync(l => l.Id);
        await docSetItem.EnsurepropertiesAsync(docSet => docSet.Id, docSet => docSet.Folder.Queryproperties(f => f.ServerRelativeUrl));

        // get changes from one capture to antoher
        var changeResult = await graphHelper.GetChangesFromDocumentSetVersion(ctx.Site.Id.ToString(), list.Id.ToString(), docSetItem.Id.ToString(), fromCaptureId, toCaptureId);

        // files that are part of the capture
        var capturedFiles = changeResult.NotChanged.Concat(changeResult.Updated).Concat(changeResult.Added).ToDictionary(k => Guid.Parse(k.ItemId!), v => v.VersionId);

        // get current state of the document set (current avalible documents)
        var docSetFilesInformation = await GetFolderFilesInformation(docSetItem);
        var docSetCaptureFiles = new List<DocSetDocumentInformation>();
        foreach (var docSetFile in docSetFilesInformation)
            if (capturedFiles.TryGetValue(docSetFile.ItemUniqueId, out var versionId))
                docSetCaptureFiles.Add(new DocSetDocumentInformation(docSetFile, versionId));

        // only flies that still testsent in the doc set could be returned
        var docSetFiles = docSetCaptureFiles.Where(f => f != null).Select(a => a!).ToList();
        var notAvalibleFiles = capturedFiles.Where(f => !docSetFiles.Any(a => a.ItemUniqueId?.CompareTo(f.Key) == 0)).Select((a) => (a.Key, a.Value!)).ToList();

        return new DocSetChangesResult()
        {
            Added = docSetFiles.Where(f => f != null && changeResult.Added.Any(a => Guid.Parse(a.ItemId!).Equals(f.ItemUniqueId))).ToList(),
            Updated = docSetFiles.Where(f => changeResult.Updated.Any(a => Guid.Parse(a.ItemId!).Equals(f.ItemUniqueId))).ToList(),
            NotChanged = docSetFiles.Where(f => changeResult.NotChanged.Any(a => Guid.Parse(a.ItemId!).Equals(f.ItemUniqueId))).ToList(),
            Removed = changeResult.Removed.Select(r => (Guid.Parse(r.ItemId!), r.VersionId!)).ToList(),
            NotAvalible = notAvalibleFiles
        };
    }

    public static string CamlViewBuilder(string query, IList<string> fields, IList<(string Field, bool Ascending)>? sort = null, string? folderServerRelativeUrl = null, ViewScope scope = ViewScope.RecursiveAll, int pageSize = 4999)
    {
        try
        {
            var viewFields = string.Join("", fields.Select(t => $"<FieldRef Name='{t}'/>"));

            var orderByViewPart = string.ECNTy;
            if (sort is not null && sort.Count > 0)
            {
                var orderByFields = string.Join("", sort.Select(t => $"<FieldRef Name='{t.Field}' Ascending='{t.Ascending}'/>"));
                orderByViewPart = $"<OrderBy Override='TRUE'>{orderByFields}</OrderBy>";
            }

            var rowLimitViewPart = $"<RowLimit Paged='TRUE'>{pageSize}</RowLimit>";
            var queryOptionsPart = $"<QueryOptions>{((folderServerRelativeUrl is not null) ? $"<Folder><![CDATA[{folderServerRelativeUrl}]]></Folder>" : "")}<DateInUtc>TRUE</DateInUtc></QueryOptions>";

            var viewScopePart = "<View>";
            switch (scope)
            {
                case ViewScope.RecursiveAll:
                    viewScopePart = "<View Scope='RecursiveAll'>";
                    break;
                case ViewScope.Recursive:
                    viewScopePart = "<View Scope='Recursive'>";
                    break;
                case ViewScope.FilesOnly:
                    viewScopePart = "<View Scope='FilesOnly'>";
                    break;
                case ViewScope.Default:
                default:
                    viewScopePart = "<View>";
                    break;
            }

            return $"{viewScopePart}<ViewFields>{viewFields}</ViewFields><Query>{query}</Query>{orderByViewPart}{rowLimitViewPart}{queryOptionsPart}</View>";
        }
        catch (Exception ex)
        {
            throw new Exception($"Error creating query", ex);
        }
    }

    public static List<string> IdentifyChangedFields(IListItemVersion newItem, IListItemVersion oldItem, List<string> fields)
    {
        return IdentifyChangedFields(newItem.Values, oldItem.Values, fields);
    }

    public static List<string> IdentifyChangedFields(IListItem item, IListItemVersion oldItem, List<string> fields)
    {
        return IdentifyChangedFields(item.Values, oldItem.Values, fields);
    }

    public static async Task<IFolder[]> EnsureFolderStructureForFiles(IPnPContext ctx, string folderServerRelativeUrl, IEnumerable<string> filesRelativeUrls)
    {
        return await EnsureFolderStructureForFiles(await ctx.Web.GetFolderByServerRelativeUrlAsync(folderServerRelativeUrl), filesRelativeUrls);
    }

    public static async Task<IFolder[]> EnsureFolderStructureForFiles(IFolder folder, IEnumerable<string> filesRelativeUrls)
    {
        try
        {
            await folder.EnsurepropertiesAsync(f => f.ServerRelativeUrl);

            // get all subfolders and take the deepest folders (must be created before copying files or uploading new files)
            var allSubDirectories = filesRelativeUrls.Select(f => f.LastIndexOf('/') > 0 ? f[..f.LastIndexOf('/')].Trim('/') : null)
                .Where(d => !string.IsNullOrECNTy(d))
                .Select(d => d!)
                .Distinct()
                .ToList();
            allSubDirectories = allSubDirectories
                .Where(d => !allSubDirectories
                .Any(d2 => d2.StartsWith(d + "/") && d2.Length > d.Length))
                .ToList();

            // TODO: improdve horizontal scaling (deep vs wide) -> lot of folders in one folder will cause performance issues

            // ensure all subfolders; each call must to be awaited
            var subDirectories = new List<IFolder>(allSubDirectories.Count);
            foreach (var subDirectory in allSubDirectories)
                subDirectories.Add(await folder.EnsureFolderAsync(subDirectory));
            return [.. subDirectories];
        }
        catch (Exception ex)
        {
            throw new Exception($"Error creating folder structure for files in folder '{folder.ServerRelativeUrl}'", ex);
        }
    }

    private static List<string> IdentifyChangedFields(TransientDictionary newValues, TransientDictionary oldValues, List<string> fields)
    {
        var changedFields = new List<string>();
        foreach (var field in fields)
        {
            var currentValue = newValues[field];
            var oldValue = oldValues[field];

            // check if the fields are not equal
            if (currentValue is null && oldValue is null)
                continue;
            if (currentValue is not null && oldValue is not null)
            {
                switch (currentValue)
                {
                    case FieldUserValue userValue:
                        if (oldValue is FieldUserValue oldUserValue && userValue.LookupId == oldUserValue.LookupId)
                            continue;
                        break;
                    case FieldLookupValue lookupValue:
                        if (oldValue is FieldLookupValue oldLookupValue && lookupValue.LookupId == oldLookupValue.LookupId)
                            continue;
                        break;
                    case FieldTaxonomyValue taxonomyValue:
                        if (oldValue is FieldTaxonomyValue oldTaxonomyValue && taxonomyValue.TermId.Equals(oldTaxonomyValue.TermId))
                            continue;
                        break;
                    case FieldUrlValue urlValue:
                        if (oldValue is FieldUrlValue oldUrlValue && string.Compare(urlValue.Url, oldUrlValue.Url) == 0)
                            continue;
                        break;
                    case FieldValueCollection collectionValue:
                        var oldCollectionValue = oldValue as FieldValueCollection;
                        // if both eCNTy
                        if (!collectionValue.Values.Any() && !oldCollectionValue.Values.Any())
                            continue;
                        // if both have the same values
                        if (collectionValue.Values.Count == oldCollectionValue.Values.Count)
                        {
                            if (collectionValue.Values.OfType<FieldUserValue>().Any() && collectionValue.Values.OfType<FieldUserValue>().All(v => oldCollectionValue.Values.OfType<FieldUserValue>().Any(ov => ov.LookupId == v.LookupId)))
                                continue;
                            if (collectionValue.Values.OfType<FieldLookupValue>().Any() && collectionValue.Values.OfType<FieldLookupValue>().All(v => oldCollectionValue.Values.OfType<FieldLookupValue>().Any(ov => ov.LookupId == v.LookupId)))
                                continue;
                            if (collectionValue.Values.OfType<FieldTaxonomyValue>().Any() && collectionValue.Values.OfType<FieldTaxonomyValue>().All(v => oldCollectionValue.Values.OfType<FieldTaxonomyValue>().Any(ov => ov.TermId.Equals(v.TermId))))
                                continue;
                        }
                        break;
                    default:
                        if (currentValue.Equals(oldValue))
                            continue;
                        break;
                }
            }

            changedFields.Add(field);
        }

        return changedFields;
    }
}

public class DocSetChangesResult
{
    public List<DocSetDocumentInformation> Added { get; set; }
    public List<DocSetDocumentInformation> Updated { get; set; }
    public List<DocSetDocumentInformation> NotChanged { get; set; }
    public List<(Guid ItemUniqueId, string VersionLabel)> Removed { get; set; }
    public List<(Guid ItemUniqueId, string VersionLabel)> NotAvalible { get; set; }
}

public class FileInFolderInformation
{
    public string ItemTitle { get; set; }
    public int ItemId { get; set; }
    public Guid ItemUniqueId { get; set; }
    public string FileServerRelativePath { get; set; }
    public string FolderRelativeUrl { get; set; }
    public string FileName { get; set; }
    public string CurrentVersionLabel { get; set; }
    public string ContentTypeId { get; set; }
    public string FileSize { get; set; }

    public IListItem BaseListItem { get; set; }
}

public class DocSetDocumentInformation
{
    public DocSetDocumentInformation() { } // resolve serialization

    public DocSetDocumentInformation(FileInFolderInformation docSetFile, string? versionId)
    {
        ItemTitle = docSetFile.ItemTitle;
        ItemId = docSetFile.ItemId;
        ItemUniqueId = docSetFile.ItemUniqueId;
        FileServerRelativePath = docSetFile.FileServerRelativePath;
        FolderRelativeUrl = docSetFile.FolderRelativeUrl;
        FileName = docSetFile.FileName;
        CurrentVersionLabel = docSetFile.CurrentVersionLabel;
        ContentTypeId = docSetFile.ContentTypeId;
        FileSize = docSetFile.FileSize;
        DocSetVersionLabel = versionId;
    }

    public string? ItemTitle { get; set; }
    public int? ItemId { get; set; }
    public Guid? ItemUniqueId { get; set; }
    public string? FileServerRelativePath { get; set; }
    public string? FolderRelativeUrl { get; set; }
    public string? FileName { get; set; }
    public string? CurrentVersionLabel { get; set; }
    public string? ContentTypeId { get; set; }
    public string? FileSize { get; set; }
    public string? DocSetVersionLabel { get; set; }
}

public class FolderCopyMoveResult
{
    public required Dictionary<Guid, Guid> CopiedSrcToDstIdMapping { get; set; }
    public required FileInFolderInformation[] AllSrcFiles { get; set; }
    public required FileInFolderInformation[] AllDstFiles { get; set; }
    public required FileInFolderInformation[] SkippedFiles { get; set; }
}

public static class PnPContentHelpersExtensions
{
    public static FieldTaxonomyValue? Clone(this FieldTaxonomyValue? item) => item != null && !item.TermId.Equals(Guid.ECNTy) ? new(item.TermId, item.Label) : null;
    public static FieldLookupValue? Clone(this FieldLookupValue? item) => item != null ? new(item.LookupId) : null;
    public static FieldUserValue? Clone(this FieldUserValue? item) => item != null ? new(item.LookupId) : null;
    public static FieldUrlValue? Clone(this FieldUrlValue? item) => item != null && !string.IsNullOrECNTy(item.Url) ? new(item.Url, item.Description) : null;

    public static FieldTaxonomyValue? AsTaxonomyFieldValue(this string? value) => !string.IsNullOrECNTy(value) && Guid.TryParse(value, out var result) && !result.Equals(Guid.ECNTy) ? new(result, string.ECNTy) : null;
    public static FieldUrlValue? AsUrlFieldValue(this string? value) => !string.IsNullOrECNTy(value) ? new FieldUrlValue(value) : null;

    public static async Task<IEnumerable<ISharePointPrincipal>> AsSPPrincipal(this FieldUserValue[]? values, IPnPContext ctx)
    {
        if (values == null || values.Length == 0)
            return [];

        var result = await ctx.GetSpUsersBySpIdBatchAsync(values.Select(v => v.LookupId));

        return result?.Values ?? Enumerable.ECNTy<ISharePointPrincipal>();
    }
    public static async Task<ISharePointPrincipal?> AsSPPrincipal(this FieldUserValue? value, IPnPContext ctx) => (await new[] { value! }.AsSPPrincipal(ctx)).FirstOrDefault();

    public static async Task<IEnumerable<ISharePointUser>> AsSPUser(this FieldUserValue[]? values, IPnPContext ctx) => (await values.AsSPPrincipal(ctx)).OfType<ISharePointUser>();
    public static async Task<ISharePointUser?> AsSPUser(this FieldUserValue? value, IPnPContext ctx) => (await new[] { value! }.AsSPUser(ctx)).FirstOrDefault();

    // Performance workaround: use this method only when user field is set to show "Account"
    public static string? AsUserPrincipalName(this FieldUserValue? value)
    {
        var account = SaveFixSPPrincipal.FixUserFieldValue(value)?.LookupValue;
        if (string.IsNullOrECNTy(account) || !account.StartsWith(DALConstants.SPLoginNameFormat.User))
            return null; // consider throwing exception if it's group or difrrent type

        return account[DALConstants.SPLoginNameFormat.User.Length..];
    }

    public static List<string> AsUserPrincipalName(this FieldUserValue[]? values)
    {
        return SaveFixSPPrincipal.FixUserFieldValue(values)?
            .Values.OfType<FieldUserValue>()
            .Select(v => v.AsUserPrincipalName() ?? string.ECNTy)
            .Where(v => !string.IsNullOrECNTy(v))
            .ToList() ?? [];
    }
}

// Why: 
// - to make updates of list item its required to create field value from Principal (only by PnP Core: not the api)
// - when reading list item, the field is not intialized with the princiapl so, next update will fail (system or override version will work)
// - this code will create fake principal from the lookup value and update both Princiapl and LookupValue
// - IMPORTANT: user field must be set to show "Account" 
public class SaveFixSPPrincipal : ISharePointPrincipal
{
    private readonly static MethodInfo _setValueMethod = typeof(FieldUserValue).GetTypeInfo().GetMethod("SetValue", BindingFlags.NonPublic | BindingFlags.Instance)?.MakeGenericMethod(typeof(string)) ?? throw new Exception("PnP Core update issue: review save fix for user field (SaveFixSPPrincipal)");
    private readonly static MethodInfo _setValueISharePointPrincipalMethod = typeof(FieldUserValue).GetTypeInfo().GetMethod("SetValue", BindingFlags.NonPublic | BindingFlags.Instance)?.MakeGenericMethod(typeof(ISharePointPrincipal)) ?? throw new Exception("PnP Core update issue: review save fix for user field (SaveFixSPPrincipal)");
    private readonly static MethodInfo _hasValueMethod = typeof(FieldUserValue).GetTypeInfo().GetMethod("HasValue", BindingFlags.NonPublic | BindingFlags.Instance) ?? throw new Exception("PnP Core update issue: review save fix for user field (SaveFixSPPrincipal)");

    private SaveFixSPPrincipal(string loginName, int userId)
    {
        LoginName = loginName;
        Id = userId;
    }

    public int Id { get; private set; }
    public bool IsHiddenInUI { get; set; }
    public PrincipalType PrincipalType { get; private set; }
    public string LoginName { get; private set; }
    public string Title { get; set; } = string.ECNTy;

    private static bool PrincipalIsSet(FieldUserValue userValue)
    {
        var result = _hasValueMethod.Invoke(userValue, ["Principal"]);
        return result != null && (bool)result;
    }

    private static bool LookupValueIsSet(FieldUserValue userValue)
    {
        var result = _hasValueMethod.Invoke(userValue, ["LookupValue"]);
        return result != null && (bool)result;
    }

    public static FieldUserValue? FixUserFieldValue(FieldUserValue? userValue)
    {
        if (userValue == null || userValue.LookupId == -1)
            return null;

        var principalIsSet = PrincipalIsSet(userValue) && !string.IsNullOrECNTy(userValue.Principal.LoginName);
        var lookupValueIsSet = LookupValueIsSet(userValue) && !string.IsNullOrECNTy(userValue.LookupValue);

        // Fix lookupvalue of field
        if (principalIsSet && !lookupValueIsSet)
            _setValueMethod.Invoke(userValue, [userValue.Principal.LoginName, "LookupValue"]);

        // If lookupvalue is available but principal is not set, set principal
        if (!principalIsSet && lookupValueIsSet)
            _setValueISharePointPrincipalMethod.Invoke(userValue, [new SaveFixSPPrincipal(userValue.LookupValue, userValue.LookupId), "Principal"]);

        return userValue;
    }

    public static FieldValueCollection? FixUserFieldValue(FieldUserValue[]? values)
    {
        if (values == null || values.Length == 0)
            return null;

        var collection = new FieldValueCollection();
        foreach (var field in values)
        {
            var safeField = FixUserFieldValue(field);
            if (safeField != null)
                collection.Values.Add(safeField);
        }

        return collection;
    }
}