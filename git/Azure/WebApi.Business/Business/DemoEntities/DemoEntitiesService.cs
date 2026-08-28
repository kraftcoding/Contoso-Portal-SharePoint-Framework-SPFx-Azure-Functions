using System.Dynamic;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Contoso.Portal.Common;


using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using Contoso.Portal.Data.DAO.DemoEntities;
using Contoso.Portal.Data.DAL.DemoEntities;
using Azure.Data.Tables;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAL;

namespace Contoso.Portal.Domains.DemoEntities;

using Contoso.Portal.Data.DAL.Helpers;

public class DemoEntitiesService(ILogger<DemoEntitiesService> logger, M365AuthHelper auth, DocumentToPdfService pdfService, IDocumentArchiveService insideService, DataStorageHelper dataStorageHelper) : ServiceBasePnP<DemoEntitiesService>(logger, auth)
{
    private readonly DocumentToPdfService _pdfService = pdfService;
    private readonly IDocumentArchiveService _insideService = insideService;
    private readonly DataStorageHelper _dataStorageHelper = dataStorageHelper;


    public async Task<byte[]> DownloadDemoEntityCertificateTemplate(IPnPContext ctx, string userUpn, string bodyId, string entidadId)
    {
        //TODO: COMprodBAR PERMISOS ESPECÍFICOS PARA REG. DE ENTIDADES DE MEMO. DEMO.
        if (!await CheckUserPermissions(ctx, userUpn))
            throw new Exception($"User '{userUpn}' does not has permissions to DownloadDemoEntityCertificateTemplate");
        try
        {
            var entidadesDAL = new EntidadesDALprovider();
            var entidad = await entidadesDAL.GetById(ctx, int.Parse(entidadId));
            if (entidad == null)
                throw new Exception($"Entidad '{entidadId}' not found on DownloadDemoEntityCertificateTemplate");

            var bodyXML = InformationToXML.RetrieveXML(entidad);
            using var templateStream = await entidadesDAL.GetTemplate(ctx);
            var template = await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, bodyXML);

            var pdfDocument = await _pdfService.ConvertToPdf(ctx, template);
            var inputDossierFile = new InputDossierFile()
            {
                Content = pdfDocument,
                CaptureDate = CultureOperations.GetCurrentSpanishTime(),
                UniqueId = Guid.NewGuid(),
                DIR3 = [string.ECNTy],
                DocumentType = DissierFileType.Other
            };
            var fileContent = await _insideService.GetContentFromDocumentToENIAsync(inputDossierFile);
            return fileContent;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(DownloadDemoEntityCertificateTemplate)} with: bodyId='{entidadId}'", ex);
        }
    }

    private async Task<bool> CheckUserPermissions(IPnPContext ctx, string upn)
    {
        var isAdmin = false;
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            var graphHelper = new GraphHelper(ctxSystem);
            var groups = await graphHelper.GetUserGroups(upn);
            isAdmin = groups.Any(group => group.DisplayName.Equals(DALConstants.DemoEntitiesGroupNames.Admin, StringComparison.OrdinalIgnoreCase) || group.DisplayName.Equals(DALConstants.DemoEntitiesGroupNames.Gestor, StringComparison.OrdinalIgnoreCase));
        });
        return isAdmin;
    }

    public async Task SyncDataToAzureDataLake(IPnPContext ctx)
    {
        await SafeExecute(() => SendEntidades(ctx), "Send Entidades to Azure Data Storage");
        await SafeExecute(() => SendScopes(ctx), "Send Scopes to Azure Data Storage");
        await SafeExecute(() => SendSections(ctx), "Send Sections to Azure Data Storage");
        await SafeExecute(() => SendCancellationReasons(ctx), "Send Causas Cancelacion to Azure Data Storage");
        await SafeExecute(() => SendDocuments(ctx), "Send Documents to Azure Data Storage");
    }

    private async Task SafeExecute(Func<Task> action, string taskName)
    {
        try
        {
            await action();
            logger.LogInformation($"Task {taskName} executed successfully at {DateTime.UtcNow}");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, $"Error on {taskName} at {DateTime.UtcNow}");
        }
    }

    private async Task SendEntidades(IPnPContext ctx)
    {
        var entidadesDAL = new EntidadesDALprovider();
        var items = await entidadesDAL.GetEntidades(ctx);

        foreach (var item in items)
        {
            var tableEntity = MapToTableEntity(item);
            await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.DemoEntities, tableEntity);
        }
    }

    private async Task SendScopes(IPnPContext ctx)
    {
        var ScopesDAL = new ScopesDALprovider();
        var Scopes = await ScopesDAL.GetAll(ctx);
        foreach (var item in Scopes)
        {
            var tableEntity = MapToTableEntity(item, DataStorage.PartitionKeys.ScopeDemoEntity);
            await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.ScopesDemoEntities, tableEntity);
        }
    }

    private async Task SendSections(IPnPContext ctx)
    {
        var SectionsDAL = new SectionsDALprovider();
        var Sections = await SectionsDAL.GetAll(ctx);
        foreach (var item in Sections)
        {
            var tableEntity = MapToTableEntity(item, DataStorage.PartitionKeys.SectionDemoEntity);
            await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.SectionsDemoEntities, tableEntity);
        }
    }

    private async Task SendCancellationReasons(IPnPContext ctx)
    {
        var causasDAL = new CancellationReasonsDALprovider();
        var causas = await causasDAL.GetAll(ctx);
        foreach (var item in causas)
        {
            var tableEntity = MapToTableEntity(item, DataStorage.PartitionKeys.CancellationReasonDemoEntity);
            await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.CancellationReasonsDemoEntities, tableEntity);
        }
    }

    private async Task SendDocuments(IPnPContext ctx)
    {
        var DocumentsDAL = new DocumentsDALprovider();
        var Documents = await DocumentsDAL.GetAll(ctx);
        foreach (var item in Documents)
        {
            var tableEntity = MapToTableEntity(item, DataStorage.PartitionKeys.DocumentDemoEntity);
            await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.DocumentsDemoEntities, tableEntity);
        }
    }

    private TableEntity MapToTableEntity(DemoEntityDAO item)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.DemoEntity, item.ID.ToString()) {
            {"Denominacion", item.Title},
            {"Folder", item.Folder_EMD},
            {"SectionId", item.Section_EMD?.LookupId},
            {"Section", item.Section_EMD?.LookupValue},
            {"ScopeTerritorialId", item.ScopeTerritorial_EMD?.LookupId},
            {"ScopeTerritorial", item.ScopeTerritorial_EMD?.LookupValue},
            {"FechaConstitucion", EnsureUtc(item.FechaConstitucion_EMD)},
            {"FechaInscripcion", EnsureUtc(item.FechaInscripcion_EMD)},
            {"DatosInscripcion", item.DatosInscripcion_EMD},
            {"DomicilioSocial", item.DomicilioSocial_EMD},
            {"NIF", item.NIF_EMD},
            {"Objeto", item.Objeto_EMD},
            {"EntidadesIntegrantes", item.EntidadesIntegrantes_EMD},
            {"IdentidadTitulares", item.IdentidadTitulares_EMD},
            {"RetestsentanteNombre", item.RetestsentanteNombre_EMD},
            {"RetestsentantePrimerApellido", item.RetestsentantePrimerApellido_EMD},
            {"RetestsentanteSegundoApellido", item.RetestsentanteSegundoApellido_EMD},
            {"RetestsentanteFechaNacimiento", EnsureUtc(item.RetestsentanteFechaNacimiento_EMD)},
            {"RetestsentanteNacionalidad", item.RetestsentanteNacionalidad_EMD},
            {"RetestsentanteTipoDocument", item.RetestsentanteTipoDocument_EMD},
            {"RetestsentanteID", item.RetestsentanteID_EMD},
            {"RetestsentanteCargo", item.RetestsentanteCargo_EMD},
            {"NumExpedienteACCEDA", item.NumExpedienteACCEDA_EMD},
            {"SolicitanteNombre", item.SolicitanteNombre_EMD},
            {"SolicitanteApellido", item.SolicitanteApellido_EMD},
            {"SolicitanteSegundoApellido", item.SolicitanteSegundoApellido_EMD},
            {"SolicitanteTipoDocument", item.SolicitanteTipoDocument_EMD},
            {"SolicitanteNumIdentificacion", item.SolicitanteNumIdentificacion_EMD},
            {"ContactoEmail", item.ContactoEmail_EMD},
            {"ContactoTelefono", item.ContactoTelefono_EMD},
            {"CancellationReason", item.CancellationReason_EMD?.LookupValue},
            {"CancellationReasonId", item.CancellationReason_EMD?.LookupId},
            {"NumExpedienteCancelacion", item.NumExpedienteCancelacion_EMD},
            {"Estado", item.Estado_EMD},
            {"Asiento", item.Asiento_EMD}
        };

        return result;
    }

    private TableEntity MapToTableEntity(TablaMaestraDAO item, string partitionKey)
    {
        var result = new TableEntity(partitionKey, item.ID.ToString()) {
            {"Titulo", item.Title},
            {"Activo", item.Activo},
            {"Descripcion", item.Descripcion ?? string.ECNTy}
        };
        return result;
    }

    private TableEntity MapToTableEntity(DocumentDAO item, string partitionKey)
    {
        var result = new TableEntity(partitionKey, item.UniqueId.ToString()) {
            {"Titulo", item.Title},
            {"TipoDocument", item.TipoDocument_EMD ?? string.ECNTy},
            {"Folder", item.Folder_EMD ?? string.ECNTy},
            {"FileRef", item.FileRef ?? string.ECNTy}
        };
        return result;
    }

    private DateTime? EnsureUtc(DateTime? date)
    {
        if (!date.HasValue)
            return null;
        return DateTime.SpecifyKind(date.Value, DateTimeKind.Utc);
    }
}