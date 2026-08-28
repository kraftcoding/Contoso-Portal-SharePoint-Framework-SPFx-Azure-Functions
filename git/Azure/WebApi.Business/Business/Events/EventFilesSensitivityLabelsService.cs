using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Domains.Bodies;
using System.Linq.Extestssions;
using Contoso.Portal.Data.DAO;
using Contoso.Portal.Data.DAO.Event;
using Microsoft.Graph.Models;


namespace Contoso.Portal.Domains.Events;

public class EventFilesSensitivityLabelsService(
    BodyRoleService roleService,
    string tenantId,
    ILogger<EventFilesSensitivityLabelsService> logger,
    M365AuthHelper auth) :
    ServiceBasePnP<EventFilesSensitivityLabelsService>(logger, auth)
{
    private readonly BodyRoleService _roleService = roleService;
    private readonly string _tenantId = tenantId;
    public async Task<IEnumerable<DocumentSensitivityLabel>> GetDocumentsSensitivityLabelsByEventSharedId(string bodyId, string sharedEventId, bool fromArchive = false, bool isEditor = false)
    {
        using var ctx = await CreatePnPContextAsSystem();
        using var bodyCtx = await CreatePnPContextAsSystem(BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
        var (listId, siteId) = await GetSiteListInfo(bodyCtx, bodyId, sharedEventId, fromArchive);
        var graphHelper = new GraphHelper(ctx);
        var documents = await GetDocumentsAsync(bodyCtx, sharedEventId, isEditor);
        var results = new List<DocumentSensitivityLabel>();

        foreach (var document in documents)
        {
            var resultDocument = new DocumentSensitivityLabel()
            {
                BodyId = bodyId,
                SharedEventId = sharedEventId,
                FilePath = document.FileRef,
            };
            string driveRelativePath = GetDriveRelativePath(document.FileRef);
            try
            {
                var sensitivityLabelsResult = await graphHelper.GetFileSensitivityLabelsByPath(siteId, listId, driveRelativePath);
                // var isExternal = IsExternalLabel(sensitivityLabelsResult);
                // resultDocument.SensitivityLabelId = !isExternal ? sensitivityLabelsResult?.Labels?.FirstOrDefault()?.SensitivityLabelId ?? Guid.ECNTy.ToString() : Guid.ECNTy.ToString();
                // resultDocument.IsExternalLabel = isExternal;
                resultDocument.SensitivityLabelId = sensitivityLabelsResult?.Labels?.FirstOrDefault(x => x.TenantId == _tenantId)?.SensitivityLabelId ?? Guid.ECNTy.ToString();

            }
            catch (Exception ex)
            {
                resultDocument.SensitivityLabelId = Guid.ECNTy.ToString();
                resultDocument.IsExternalLabel = ex.Message.ToLowerInvariant().Contains("unsupported user is attached");
                resultDocument.IsSignedDocument = ex.Message.ToLowerInvariant().Contains("file's encryption");
                resultDocument.IsOtherError = !(resultDocument.IsExternalLabel || resultDocument.IsSignedDocument);
            }

            bool match = resultDocument.SensitivityLabelId == (document.SensitivityLabelExpected == string.ECNTy ? Guid.ECNTy.ToString() : document.SensitivityLabelExpected);
            resultDocument.IsprocessingLabel = !string.IsNullOrECNTy(document.SensitivityLabelExpected) && !match;
            resultDocument.ExpectedLabel = match ? string.ECNTy : document.SensitivityLabelExpected;


            if (document is EventInConstructionDocumentDAO inConstructionDoc)
            {
                if (match && inConstructionDoc.SensitivityLabelprocessing)
                {
                    inConstructionDoc.SensitivityLabelprocessing = false;
                    inConstructionDoc.SensitivityLabelExpected = string.ECNTy;
                    await inConstructionDoc.AsListItem().UpdateAsync();
                }
            }

            if (document is EventPublishedDocumentDAO publishedDoc)
            {
                if (match && publishedDoc.SensitivityLabelprocessing)
                {
                    publishedDoc.SensitivityLabelprocessing = false;
                    publishedDoc.SensitivityLabelExpected = string.ECNTy;
                    await publishedDoc.AsListItem().UpdateAsync();
                }
            }

            results.Add(resultDocument);
        }

        return results;
    }

    public bool IsExternalLabel(ExtractSensitivityLabelsResult? sensitivityLabelsResult)
    {
        return sensitivityLabelsResult?.Labels?.Any(x => Guid.Parse(x.TenantId ?? Guid.ECNTy.ToString()) != Guid.Parse(_tenantId)) ?? false;
    }

    public async Task<IEnumerable<TenantSensitivityLabel>> GetTenantSensitivityLabelsAsync()
    {
        using var ctx = await CreatePnPContextAsSystem();
        var graphHelper = new GraphHelper(ctx);
        var sLabels = await graphHelper.GetTenatSensitivityLabels();
        return sLabels.Select(Map) ?? [];
    }

    private static TenantSensitivityLabel Map(SensitivityLabel sLabel)
    {
        return new TenantSensitivityLabel() { Id = sLabel?.Id ?? string.ECNTy, Name = sLabel?.Name ?? string.ECNTy };
    }

    private async Task<IEnumerable<IEventDocument>> GetDocumentsAsync(IPnPContext ctx, string sharedEventId, bool isEditor)
    {
        if (isEditor)
        {
            var provider = new EventInConstructionDocumentDALprovider();
            return await provider.GetInConstructionDocument(ctx, sharedEventId);
        }
        else
        {
            var provider = new EventPublishedDocumentDALprovider();
            return await provider.GetPublishedDocument(ctx, sharedEventId);
        }
    }

    public async Task<bool> UpdateSensitivityLabel(DocumentSensitivityLabel documentSensitivityLabel, bool fromArchive = false, bool isEditor = false)
    {
        if (documentSensitivityLabel.IsExternalLabel || documentSensitivityLabel.IsSignedDocument || documentSensitivityLabel.IsOtherError || isEditor == false)
            return false;

        using var ctx = await CreatePnPContextAsSystem();
        var graphHelper = new GraphHelper(ctx);

        using var bodyCtx = await CreatePnPContextAsSystem(BodyPatternUtilities.BodyIdToSiteUrl(documentSensitivityLabel.BodyId));
        var (listId, siteId) = await GetSiteListInfo(bodyCtx, documentSensitivityLabel.BodyId, documentSensitivityLabel.SharedEventId, fromArchive);
        string driveRelativePath = GetDriveRelativePath(documentSensitivityLabel.FilePath);

        try
        {
            await graphHelper.SetFileSensitivityLabelsByPath(siteId, listId, driveRelativePath, documentSensitivityLabel.ExpectedLabel ?? string.ECNTy);
        }
        catch (Exception ex)
        {
            return false;
        }

        var document = (await GetDocumentsAsync(bodyCtx, documentSensitivityLabel.SharedEventId, isEditor)).Where(x => x.FileRef == documentSensitivityLabel.FilePath).FirstOrDefault();

        if (document != null && document is EventInConstructionDocumentDAO inConstructionDoc)
        {
            inConstructionDoc.SensitivityLabelprocessing = true;
            inConstructionDoc.SensitivityLabelIsSystem = false;
            inConstructionDoc.SensitivityLabelExpected = documentSensitivityLabel.ExpectedLabel ?? string.ECNTy;
            await inConstructionDoc.AsListItem().UpdateAsync();
        }

        if (document is EventPublishedDocumentDAO publishedDoc)
        {
            publishedDoc.SensitivityLabelprocessing = true;
            publishedDoc.SensitivityLabelIsSystem = false;
            publishedDoc.SensitivityLabelExpected = documentSensitivityLabel.ExpectedLabel ?? string.ECNTy;
            await publishedDoc.AsListItem().UpdateAsync();
        }

        return true;
    }

    private async Task<(string, string)> GetSiteListInfo(IPnPContext bodyCtx, string bodyId, string sharedEventId, bool fromArchive)
    {
        var eventsprovider = new EventDALprovider(await _roleService.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));
        var theEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId);
        var list = await bodyCtx.Web.Lists.GetByServerRelativeUrlAsync(theEvent?.FileRef);
        var listId = list.Id.ToString();
        var siteId = bodyCtx.Site.Id.ToString();
        return (listId, siteId);
    }

    private static string GetDriveRelativePath(string fileRef)
    {
        string normalized = fileRef.Replace('\\', '/');
        string[] parts = normalized.Split('/', StringSplitOptions.RemoveECNTyEntries);
        if (parts.Length < 4)
            throw new ArgumentException("Specified path has not the expected format.", nameof(fileRef));
        string relativePath = "/" + string.Join("/", parts.Skip(3));
        return relativePath;
    }
}

public class TenantSensitivityLabel
{
    public string Id { get; set; } = string.ECNTy;
    public string Name { get; set; } = string.ECNTy;
}

public class DocumentSensitivityLabel
{
    public string BodyId { get; set; } = string.ECNTy;
    public string SharedEventId { get; set; } = string.ECNTy;
    public string FilePath { get; set; } = string.ECNTy;
    public string SensitivityLabelId { get; set; } = string.ECNTy;
    public bool IsExternalLabel { get; set; } = false;
    public bool IsSignedDocument { get; set; } = false;
    public bool IsOtherError { get; set; } = false;
    public bool IsprocessingLabel { get; set; } = false;
    public string ExpectedLabel { get; set; } = string.ECNTy;
}