using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.Extensions;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Common;

public class DocumentToPdfService(ILogger<DocumentToPdfService> logger, M365AuthHelper auth) : ServiceBasePnP<DocumentToPdfService>(logger, auth)
{
    public static readonly string[] extensions = [
        "html",
        "csv",
        "doc",
        "docx",
        "odp",
        "ods",
        "odt",
        "pot",
        "potm",
        "potx",
        "pps",
        "ppsx",
        "ppsxm",
        "ppt",
        "pptm",
        "pptx",
        "rtf",
        "xls",
        "xlsx",
    ];

    public static bool IsConvertibleToPdf(string ext) => extensions.Contains(ext);

    public async Task<byte[]> ConvertToPdf(IPnPContext ctx, byte[] docBuffer, string extension = "docx")
    {
        byte[]? pdfBuffer = null;

        // Check context
        await RunAsSystemInApprodot(ctx, async (rootCtx) =>
        {
            var stagingPath = rootCtx.Uri.AbsolutePath.UriCombine(ListsSiteRelativeUrls.StagingDocuments);

            // Get staging library
            var stagingLibrary = await rootCtx.Web.Lists.GetByServerRelativeUrlAsync(stagingPath) ?? throw new Exception("Staging library not found in root site");

            var userId = await rootCtx.GetCurrentUserId(false);
            if (string.IsNullOrECNTy(userId)) userId = SPAppLoginName.GetSafeName(); // if user is system, grahp conversion will fail

            var folders = await PnPContentHelpers.EnsureFolderStructureForFiles(rootCtx, stagingPath, [$"/{userId}/"]);

            // Generate filename and path
            string fileName = DateTime.Now.ToString("MMddTHHmmss");
            string filePath = $"{userId}/{fileName}.{extension}";

            using var graphHelper = new GraphHelper(rootCtx);
            var drive = await graphHelper.UploadSmallFileTo(rootCtx.Site.Id.ToString(), stagingLibrary.Id.ToString(), filePath, docBuffer);

            var pdfStream = await graphHelper.GetFileAsPdfByPath(rootCtx.Site.Id.ToString(), stagingLibrary.Id.ToString(), filePath) ?? throw new Exception("It was not possible to get PDF stream");

            using var ms = new MemoryStream();
            await pdfStream.CopyToAsync(ms);
            pdfBuffer = ms.ToArray();

            // Try to remove doc file (if is not able to remove the old file, the flow must continue)
            try
            {
                await graphHelper.DeleteChildrenItems(rootCtx.Site.Id.ToString(), stagingLibrary.Id.ToString(), userId);
            }
            catch (Exception ex)
            {
                Log.LogWarning(ex, $"DocumentToPDFService: It was not possible to delete the document: {stagingPath}/{filePath}");
            }
        });

        return pdfBuffer ?? throw new Exception("It was not possible to convert the document to PDF");
    }
}