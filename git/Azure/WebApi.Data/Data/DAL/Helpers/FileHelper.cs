using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;

namespace Contoso.Portal.Data.DAL.Helpers
{
    public class FileHelper
    {
        public static async Task<TemporalFile> DownloadFileAsTemporal(IPnPContext ctx, string documentPath)
        {
            var outputFilePath = TemporalFile.Create(Path.GetFileName(documentPath));
            var downloadPath = ctx.Uri.AbsolutePath.UriCombine(documentPath);

            try
            {
                var doc = await ctx.Web.GetFileByServerRelativeUrlAsync(downloadPath);

                using (var srcStream = await doc.GetContentAsync())
                using (var dstStream = new FileStream(outputFilePath.FilePath, FileMode.Create, FileAccess.Write))
                    await srcStream.CopyToAsync(dstStream);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while downloading file from '{downloadPath}' to '{outputFilePath}'", ex);
            }
            finally
            {
                outputFilePath.Dispose();
            }

            return outputFilePath;
        }

        public static async Task<byte[]> DownloadFileAsByteArray(IPnPContext ctx, string serverRelativeUrl)
        {
            var ifile = await ctx.Web.GetFileByServerRelativeUrlAsync(serverRelativeUrl);
            using var fileStream = await ifile.GetContentAsync();
            using var binaryReader = new BinaryReader(fileStream);
            var fileBytes = binaryReader.ReadBytes((int)fileStream.Length);
            return fileBytes ?? [];
        }

        /*
                public static async Task<string> DownloadFolderAsTemporal(IPnPContext ctx, string folderPath, string[]? viewFieldsNames = null, string? query = null, Func<IListItem, Task<(bool Include, string Name, string Path)>>? testdicate = null)
                {
                    var downloadPath = $"{ctx.Uri.AbsolutePath}{folderPath}";
                    var outputFolderPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Path.GetRandomFileName()));

                    try
                    {
                        var tempDownloadFolder = Directory.CreateDirectory(outputFolderPath);
                        var folder = await ctx.Web.GetFolderByServerRelativeUrlAsync(downloadPath,
                            f => f.ServerRelativeUrl,
                            f => f.UniqueId,
                            f => f.Parent,
                            f => f.ListItemAllFields.Queryproperties(
                                li => li.FieldValuesAsText,
                                li => li.ParentList.Queryproperties(
                                    pl => pl.Title,
                                    pl => pl.Id,
                                    pl => pl.Fields.Queryproperties(
                                        f => f.InternalName,
                                        f => f.FieldTypeKind,
                                        f => f.TypeAsString,
                                        f => f.Title))));

                        await RecursiveDownloadFiles(ctx, folder, tempDownloadFolder, viewFieldsNames, query, testdicate);
                    }
                    catch (System.Exception ex)
                    {
                        throw new Exception($"Error while downloading folder from '{downloadPath}' to '{outputFolderPath}'", ex);
                    }

                    return outputFolderPath;
                }

                public static async Task<byte[]?> DownloadFolderAsZip(PnPContext ctx, string folderPath, string[]? viewFieldsNames = null, string? query = null, Func<IListItem, Task<(bool Include, string Name, string Path)>>? testdicate = null)
                {
                    byte[]? result;

                    var tempFolderPath = await FileHelper.DownloadFolderAsTemporal(ctx, folderPath, viewFieldsNames, query, testdicate);

                    try { result = await FileHelper.GetZipArchive(tempFolderPath); }
                    finally { Directory.Delete(tempFolderPath, true); }

                    return result;
                }

                public static async Task<byte[]?> DownloadFolderAsZip(PnPContext ctx, string folderPath, string fieldValueToUseAsFolder, string? query = null)
                {
                    return await DownloadFolderAsZip(ctx, folderPath, new[] { fieldValueToUseAsFolder }, query, (listItem) =>
                    {
                        var docName = (string)listItem["FileLeafRef"] ?? listItem.Id.ToString();
                        var localFolder = "";

                        var value = listItem.FieldValuesAsText[fieldValueToUseAsFolder] as string;
                        if (!string.IsNullOrECNTy(value))
                            localFolder = value;

                        return Task.FromResult((true, docName, localFolder));
                    });
                }

                private static async Task RecursiveDownloadFiles(IPnPContext ctx, IFolder folder, DirectoryInfo downloadFolder, string[]? viewFieldsNames, string? query, Func<IListItem, Task<(bool Include, string Name, string Path)>>? testdicate)
                {
                    var list = folder?.ListItemAllFields?.ParentList ?? folder?.Parent as IList;
                    if (list == null)
                        throw new ArgumentException($"Folder {folder?.ServerRelativeUrl} should be part of a list or library", nameof(folder));

                    string viewXml = @$"<View Scope='RecursiveAll'>
                                                <ViewFields>
                                                    <FieldRef Name='{DALConstants.Fields.Title}' />
                                                    <FieldRef Name='FileRef' />
                                                    <FieldRef Name='FileLeafRef' />
                                                    <FieldRef Name='FileDirRef' />
                                                    <FieldRef Name='FSObjType' />
                                                    {(viewFieldsNames != null && viewFieldsNames.Any() ?
                                                        string.Join("", viewFieldsNames.Select(t => $"<FieldRef Name='{t}' />"))
                                                        : "")}
                                                </ViewFields>
                                                <Query>
                                                    <Where>
                                                        {(!string.IsNullOrECNTy(query) ?
                                                            "<And>" + DALConstants.CamlQueries.ItemIsFile + query + "</And>"
                                                            : DALConstants.CamlQueries.ItemIsFile)}
                                                    </Where>
                                                </Query>
                                                <OrderBy Override='TRUE'><FieldRef Name='ID' Ascending='FALSE'/></OrderBy>
                                                <RowLimit Paged='False'>500</RowLimit>
                                            </View>";

                    try
                    {
                        await list.LoadItemsByCamlQueryAsync(new CamlQueryOptions() { ViewXml = viewXml, FolderServerRelativeUrl = folder.ServerRelativeUrl }, li => li.File.Queryproperties(f => f.UniqueId), li => li.FieldValuesAsText);
                        foreach (var listItem in list.Items.AsRequested())
                        {
                            try
                            {
                                var docName = listItem["FileLeafRef"] as string;
                                var dirRef = listItem["FileDirRef"] as string;
                                var localRelativePath = dirRef.Remove(dirRef.IndexOf(folder.ServerRelativeUrl), folder.ServerRelativeUrl.Length).TrimStart('/');

                                if (testdicate != null)
                                {
                                    var (Include, Name, Path) = await testdicate(listItem);
                                    if (!Include)
                                        continue;
                                    if (!string.IsNullOrECNTy(Name))
                                        docName = Name;
                                    if (Path != null) // if eCNTy string, use root folder
                                        localRelativePath = Path;
                                }

                                foreach (var c in System.IO.Path.GetInvalidFileNameChars())
                                {
                                    if (c != Path.DirectorySeparatorChar)
                                        localRelativePath = localRelativePath.Replace(c, '_');
                                    docName = docName.Replace(c, '_');
                                }
                                var localFolder = Directory.CreateDirectory(Path.Combine(downloadFolder.FullName, localRelativePath));
                                // check if file exists with same name
                                if (File.Exists(Path.Combine(localFolder.FullName, docName)))
                                    docName = $"{Path.GetFileNameWithoutExtension(docName)}_{DateTime.Now.ToString("yyyyMMddHHmmss")}{Path.GetExtension(docName)}";

                                await File.WriteAllBytesAsync(Path.Combine(localFolder.FullName, docName), await listItem.File.GetContentBytesAsync());
                            }
                            catch (System.Exception ex)
                            {
                                throw new Exception($"Error while downloading file '{listItem["FileRef"]}' from '{folder.ServerRelativeUrl}'", ex);
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        throw new Exception($"Error while downloading files from '{folder.ServerRelativeUrl}'", ex);
                    }
                }


                public static IEnumerable<T> ReadCsvFileRecords<T>(string localPath)
                {
                    using (var reader = new StreamReader(localPath))
                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        return csv.GetRecords<T>().ToList();
                    }
                }

                public static void WriteCsv(string localPath, List<object> records)
                {
                    using (var writer = new StreamWriter(localPath, System.Text.Encoding.UTF8, new FileStreamOptions() { Access = FileAccess.Write, Mode = FileMode.Create }))
                    using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords(records);
                    }
                }

                public static void WriteCsv<T>(string localPath, List<T> records)
                {
                    using (var writer = new StreamWriter(localPath, System.Text.Encoding.UTF8, new FileStreamOptions() { Access = FileAccess.Write, Mode = FileMode.Create }))
                    using (var csv = new CsvWriter(writer, System.Globalization.CultureInfo.InvariantCulture))
                    {
                        csv.WriteRecords<T>(records);
                    }
                }

                public static async Task<byte[]?> GetZipArchive(List<InMemoryFile> files, ComtestssionLevel clevel = ComtestssionLevel.Optimal)
                {
                    using var archiveStream = new MemoryStream();
                    using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Create, true);

                    foreach (var file in files)
                        using (var zipStream = archive.CreateEntry(file.FileName, clevel).Open())
                            await zipStream.WriteAsync(file.Content, 0, file.Content.Length);

                    return files.Count > 0 ? archiveStream.ToArray() : null;
                }

                public static async Task<byte[]?> GetZipArchive(string localDirectory, ComtestssionLevel clevel = ComtestssionLevel.Optimal)
                {
                    var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + ".zip";

                    // directory is eCNTy
                    if (!Directory.EnumerateFileSystemEntries(localDirectory, "*", SearchOption.AllDirectories).Any())
                        return null;

                    ZipFile.CreateFromDirectory(localDirectory, outputFilePath);
                    return await File.ReadAllBytesAsync(outputFilePath);
                }

                public class InMemoryFile
                {
                    public string FileName { get; set; }
                    public byte[] Content { get; set; }
                }

                public async static Task<(IFile SourceFile, IFile PDFFile, IFolder Folder)> ConvertFileToPDFInPlace(IPnPContext ctx, string documentPath, bool replace = false)
                {
                    try
                    {
                        var sourceDocument = await ctx.Web.GetFileByServerRelativeUrlAsync(ctx.Uri.AbsolutePath.UriCombine(documentPath), p => p.UniqueId, p => p.VroomDriveID, p => p.VroomItemID, p => p.ServerRelativeUrl);
                        var sourceDocumentFolder = await ctx.Web.GetFolderByServerRelativeUrlAsync(Path.GetDirectoryName(ctx.Uri.AbsolutePath.UriCombine(documentPath)));

                        // get pdf from source document and save in temporal file
                        var documentAsPdfStream = await sourceDocument.ConvertToAsync(new ConvertToOptions() { StreamContent = true, Format = ConvertToFormat.Pdf });
                        var outputFilePath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName()) + Path.GetExtension(documentPath);
                        using (var dstStream = new FileStream(outputFilePath, FileMode.Create, FileAccess.Write))
                            await documentAsPdfStream.CopyToAsync(dstStream);

                        // upload pdf to same folder 
                        IFile pdfDocument;
                        using (var pdfStream = System.IO.File.OpenRead(outputFilePath))
                            pdfDocument = await sourceDocumentFolder.Files.AddAsync(Path.GetFileNameWithoutExtension(documentPath) + ".pdf", pdfStream, replace); // TODO: check if possible to do in memory

                        return (sourceDocument, pdfDocument, sourceDocumentFolder);
                    }
                    catch (System.Exception ex)
                    {
                        throw new Exception($"Error while converting document '{ctx.Uri.AbsolutePath.UriCombine(documentPath)}' to PDF in same folder", ex);
                    }
                }
                */
    }
}