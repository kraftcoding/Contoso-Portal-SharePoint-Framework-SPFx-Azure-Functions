using System.Text.Json;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAL.Management;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.DAO.Management;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;

using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Domains.Events;
using Azure.Data.Tables;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Data.DTO.Tasks;
using System.Text;
using Contoso.Portal.Data.DAL;

namespace Contoso.Portal.Domains.PowerBiReports;

public class PowerBiReportsService(ILogger<PowerBiReportsService> logger, M365AuthHelper auth, DataStorageHelper dataStorageHelper, BodyRoleService roleSerivce, ConfigDepartmentsService ConfigDepartmentsService, EventMinutesService eventMinutesService, EventAttendanceService eventAttendanceService) : ServiceBasePnP<PowerBiReportsService>(logger, auth)
{
    private readonly BodyRoleService _roleService = roleSerivce;
    private readonly ConfigDepartmentsService _ConfigDepartmentsService = ConfigDepartmentsService;
    private readonly EventMinutesService _eventMinutesService = eventMinutesService;
    private readonly EventAttendanceService _eventAttendanceService = eventAttendanceService;
    private readonly DataStorageHelper _dataStorageHelper = dataStorageHelper;
    private readonly string _landingSite = "Contoso";
    public DateTime _currentReportRunDate;
    public DateTime? _latestReportRunDate = null;
    private StringBuilder stringBuilder = new();
    public async Task RunReport(IPnPContext ctx, string? latestReportRunDate)
    {
        stringBuilder = new();
        _currentReportRunDate = DateTime.UtcNow;
        var typeOfReport = string.IsNullOrECNTy(latestReportRunDate) ? Management.TypeOfReport.Full : Management.TypeOfReport.Incremental;
        var requestRowId = await AddPowerBiprocessRequest(ctx, typeOfReport);

        if (typeOfReport == Management.TypeOfReport.Full)
        {
            await SafeExecute(DeleteAllRecords, nameof(DeleteAllRecords));
        }
        else
        {
            _latestReportRunDate = Convert.ToDateTime(latestReportRunDate).ToUniversalTime();
        }

        var timePeriod = typeOfReport == Management.TypeOfReport.Full ? _currentReportRunDate + "" : $"period between {_latestReportRunDate} and {_currentReportRunDate}";

        stringBuilder.AppendLine($"Starting data dump process from {timePeriod}, at {DateTime.UtcNow}");

        var tasks = new List<Task>()
        {
            SafeExecute(processEXTERNALBodies, nameof(processEXTERNALBodies)),
            SafeExecute(() => processBusinessUnit(ctx), nameof(processBusinessUnit)),
            SafeExecute(processContosoBodies, nameof(processContosoBodies))
        };

        await Task.WhenAll(tasks);
        stringBuilder.AppendLine($"Finished data dump process from {timePeriod}, at {DateTime.UtcNow}");


        await SaveLogsFile(ctx, $"ReportLog{_currentReportRunDate:yyyy_MM_dd_HH_mm_ss}.txt", stringBuilder.ToString(), new Dictionary<string, object>
        {
            { "Title", $"ReportLog{_currentReportRunDate:yyyy_MM_dd_HH_mm_ss}.txt" },
            { "Tipoprodceso", typeOfReport },
            { "EstadoPeticion", !stringBuilder.ToString().Contains("Error") ? "OK" : "Error" },
            { "Inicioprodceso", _currentReportRunDate },
            { "Finprodceso", DateTime.UtcNow },
        });

        await DeletePowerBiprocessRequest(ctx, requestRowId);

    }

    #region Data prodvisioning
    private async Task DeleteAllRecords()
    {
        List<string> tables = [DataStorage.TableNames.DepartmentsEXTERNAL, DataStorage.TableNames.RelacionBusinessAreaDivision, DataStorage.TableNames.DepartmentsContoso, DataStorage.TableNames.Meetings, DataStorage.TableNames.Minutes, DataStorage.TableNames.DocumentsEvento, DataStorage.TableNames.TasksMinutesApprodval, DataStorage.TableNames.TasksAttendance, DataStorage.TableNames.TasksCertification, DataStorage.TableNames.TasksDelegation, DataStorage.TableNames.TasksMinutesModification, DataStorage.TableNames.AgendaItems, DataStorage.TableNames.Votes];
        foreach (var table in tables)
        {
            await _dataStorageHelper.DeleteAllEntitiesAsync(table);
        }

    }

    private async Task processEXTERNALBodies()
    {
        stringBuilder.AppendLine($"Starting processing EXTERNAL bodies at {DateTime.UtcNow}");

        //ÓRGANOS DE EXTERNAL
        var EXTERNALBodies = await _ConfigDepartmentsService.GetEXTERNALConfigDepartments();
        if (_latestReportRunDate.HasValue)
        {
            EXTERNALBodies = EXTERNALBodies.Where(x => x.Modified.ToUniversalTime() >= _latestReportRunDate && x.Modified.ToUniversalTime() <= _currentReportRunDate);
        }
        await _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.DepartmentsEXTERNAL, EXTERNALBodies.Select(Map).ToList());

        stringBuilder.AppendLine($"Completed processing EXTERNAL bodies at {DateTime.UtcNow}");
    }

    private async Task processBusinessUnit(IPnPContext ctx)
    {
        stringBuilder.AppendLine($"Starting processing BusinessArea-Ministry relations at {DateTime.UtcNow}");

        //RELACIÓN ÁREA SECTORIAL - Division
        var BusinessUnit = await new BusinessUnitDALprovider().GetBusinessAreasMinistries(ctx);
        if (_latestReportRunDate.HasValue)
        {
            BusinessUnit = BusinessUnit.Where(x => x.Modified.ToUniversalTime() >= _latestReportRunDate && x.Modified.ToUniversalTime() <= _currentReportRunDate);
        }
        await _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.RelacionBusinessAreaDivision, BusinessUnit.Select(Map).ToList());

        stringBuilder.AppendLine($"Completed processing BusinessArea-Ministry relations at {DateTime.UtcNow}");
    }

    private async Task processContosoBodies()
    {
        stringBuilder.AppendLine($"Starting processing Contoso bodies at {DateTime.UtcNow}");

        //ÓRGANOS DE Contoso
        var bodies = await _ConfigDepartmentsService.GetConfigDepartments();
        foreach (var body in bodies)
        {
            stringBuilder.AppendLine($"processing body {body.Title} at {DateTime.UtcNow}");

            if (!_latestReportRunDate.HasValue || (body.Modified.ToUniversalTime() >= _latestReportRunDate && body.Modified.ToUniversalTime() <= _currentReportRunDate))
            {
                await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.DepartmentsContoso, Map(body));
            }

            IPnPContext bodyCtx;
            try
            {
                bodyCtx = await CreatePnPContextAsSystem(BodyPatternUtilities.BodyIdToSiteUrl(body.Title));
            }
            catch (Exception)
            {
                stringBuilder.AppendLine($"Warning - the body {body.Title} has not been created correctly");
                stringBuilder.AppendLine($"Finished processing body {body.Title} at {DateTime.UtcNow}");
                continue;
            }
            await SafeExecute(() => processBodyEvents(bodyCtx, body.Title), nameof(processBodyEvents));
            stringBuilder.AppendLine($"Finished processing body {body.Title} at {DateTime.UtcNow}");

        }

        stringBuilder.AppendLine($"Completed processing Contoso bodies at {DateTime.UtcNow}");
    }

    private async Task processBodyEvents(IPnPContext bodyCtx, string bodyTitle)
    {
        stringBuilder.AppendLine($"Starting processing Events at {DateTime.UtcNow}");

        //Meetings
        var eventsprovider = new EventDALprovider(await _roleService.GetDefaultEventSourceByBodyRole(bodyCtx, bodyTitle, false));
        var eventsproviderArchived = new EventDALprovider(await _roleService.GetDefaultEventSourceByBodyRole(bodyCtx, bodyTitle, true));

        var bodyEvents = await eventsprovider.GetByView(bodyCtx, eventsprovider.DefaultView);
        var archivedBodyEvents = await eventsproviderArchived.GetByView(bodyCtx, eventsproviderArchived.DefaultView);

        var allEvents = bodyEvents.Concat(archivedBodyEvents);

        foreach (var bodyEvent in allEvents)
        {
            stringBuilder.AppendLine($"processing event {bodyEvent.IdMeeting}");
            if (!_latestReportRunDate.HasValue || (bodyEvent.Modified.ToUniversalTime() >= _latestReportRunDate && bodyEvent.Modified.ToUniversalTime() <= _currentReportRunDate))
            {
                await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.Meetings, Map(bodyEvent, bodyTitle));
            }
            bool isArchived = archivedBodyEvents.Any(x => x.IdMeeting == bodyEvent.IdMeeting);
            await SafeExecute(() => processEventDetails(bodyCtx, bodyTitle, bodyEvent, isArchived), nameof(processEventDetails));
            stringBuilder.AppendLine($"Finished processing event {bodyEvent.IdMeeting}");

        }

        stringBuilder.AppendLine($"Completed processing Events at {DateTime.UtcNow}");
    }

    private async Task processEventDetails(IPnPContext bodyCtx, string bodyTitle, BaseEventDAO bodyEvent, bool isArchived)
    {
        await SafeExecute(() => processAgendaItems(bodyCtx, bodyTitle, bodyEvent.IdMeeting, isArchived), nameof(processAgendaItems));
        await SafeExecute(() => processVotations(bodyCtx, bodyTitle, bodyEvent.IdMeeting, isArchived), nameof(processVotations));
        await SafeExecute(() => processMinutesInformation(bodyCtx, bodyTitle, bodyEvent.IdMeeting, isArchived), nameof(processMinutesInformation));
        await SafeExecute(() => processDocuments(bodyCtx, bodyEvent.IdMeeting), nameof(processDocuments));
        await SafeExecute(() => processTasks(bodyCtx, bodyEvent.IdMeeting, isArchived), nameof(processTasks));
    }

    private async Task processAgendaItems(IPnPContext bodyCtx, string bodyTitle, string idMeeting, bool isArchived)
    {
        //ÓRDENES DEL DÍA
        stringBuilder.AppendLine($"Starting processing Agenda Items at {DateTime.UtcNow}");
        var eventsprovider = new EventDALprovider(await _roleService.GetDefaultEventSourceByBodyRole(bodyCtx, bodyTitle, isArchived));
        var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);
        var (_, agendaItems) = await agendaItemsprovider.GetAllDTOByEvent(bodyCtx, idMeeting, true);
        foreach (var agendaItem in agendaItems)
        {
            if (!_latestReportRunDate.HasValue || (agendaItem.ItemDAO.Modified.ToUniversalTime() >= _latestReportRunDate && agendaItem.ItemDAO.Modified.ToUniversalTime() <= _currentReportRunDate))
            {
                await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.AgendaItems, Map(agendaItem));
            }
        }

        stringBuilder.AppendLine($"Completed processing Agenda Items at {DateTime.UtcNow}");
    }

    private async Task processVotations(IPnPContext bodyCtx, string bodyTitle, string idMeeting, bool isArchived)
    {
        //Votes
        stringBuilder.AppendLine($"Starting processing Votation Details at {DateTime.UtcNow}");
        var eventsprovider = new EventDALprovider(await _roleService.GetDefaultEventSourceByBodyRole(bodyCtx, bodyTitle, isArchived));
        var votationDetailDALprovider = new EventVotationDetailDALprovider(eventsprovider);
        await RunAsSystem(bodyCtx, async (ctxSystem) =>
        {
            var (_, votations) = await votationDetailDALprovider.GetAllDTOByEvent(ctxSystem, idMeeting);
            if (_latestReportRunDate.HasValue)
            {
                votations = votations.Where(x => x.ItemDAO.Modified.ToUniversalTime() >= _latestReportRunDate && x.ItemDAO.Modified.ToUniversalTime() <= _currentReportRunDate);
            }

            foreach (var votation in votations)
            {
                await _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.Votes, Map(votation));
            }
        });
        stringBuilder.AppendLine($"Completed processing Votation Details at {DateTime.UtcNow}");
    }

    private async Task processMinutesInformation(IPnPContext bodyCtx, string bodyTitle, string idMeeting, bool isArchived)
    {

        stringBuilder.AppendLine($"Starting processing Minutes Information at {DateTime.UtcNow}");

        //Minutes
        var eventMinutes = await _eventMinutesService.GetMinutesInformationByEvent(bodyCtx, bodyTitle, idMeeting, isArchived);
        await _dataStorageHelper.UpsertEntityAsync(DataStorage.TableNames.Minutes, Map(eventMinutes, idMeeting));

        stringBuilder.AppendLine($"Completed processing Minutes Information at {DateTime.UtcNow}");


    }

    private async Task processDocuments(IPnPContext bodyCtx, string idMeeting)
    {

        stringBuilder.AppendLine($"Starting processing Documents relations at {DateTime.UtcNow}");

        //Documents
        var documentDalprovider = new EventPublishedDocumentDALprovider();
        var documents = await documentDalprovider.GetPublishedDocument(bodyCtx, idMeeting);
        if (_latestReportRunDate.HasValue)
        {
            documents = documents.Where(x => x.Modified.ToUniversalTime() >= _latestReportRunDate && x.Modified.ToUniversalTime() <= _currentReportRunDate);
        }
        await _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.DocumentsEvento, documents.Select(Map).ToList());

        stringBuilder.AppendLine($"Completed processing Documents at {DateTime.UtcNow}");

    }

    public async Task processTasks(IPnPContext bodyCtx, string idMeeting, bool isArchived)
    {

        stringBuilder.AppendLine($"Starting processing Tasks at {DateTime.UtcNow}");

        //var Request = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.Request)).ToList();
        var userTasksprovider = new TasksDALprovider(isArchived ? Source.Archived : Source.InConstruction);
        var userTasks = await userTasksprovider.GetFilterTaskByEvent(bodyCtx, idMeeting);
        if (_latestReportRunDate.HasValue)
        {
            userTasks = userTasks.Where(x => x.ItemDAO.Modified.ToUniversalTime() >= _latestReportRunDate && x.ItemDAO.Modified.ToUniversalTime() <= _currentReportRunDate);
        }

        //TAREAS DE AprodBACIÓN DE Minutes
        var approdvalMinutesTasks = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestAprodbacionMinutes));

        //TAREAS DE Attendance
        var taskAttendance = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestAttendance)).ToList();

        //TAREAS DE CERTIFICACIÓN DE Minutes
        var certificationTasks = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestCertificacion)).ToList();

        //TAREAS DE Delegation
        var delegateTasks = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestDelegation)).ToList();

        //TAREAS DE MODIFICACIÓN DE Minutes
        var modificationMinutesTasks = userTasks.Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestModificacion)).ToList();

        var tasks = new List<Task>() {
            SafeExecute(() => _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.TasksMinutesApprodval, approdvalMinutesTasks.Select(Map).ToList()), nameof(processTasks) + DataStorage.PartitionKeys.TaskMinutes),
            SafeExecute(() => _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.TasksAttendance, taskAttendance.Select(Map).ToList()), nameof(processTasks) + DataStorage.PartitionKeys.TaskAttendance),
            SafeExecute(() => _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.TasksCertification, certificationTasks.Select(Map).ToList()), nameof(processTasks) + DataStorage.PartitionKeys.TaskCertification),
            SafeExecute(() => _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.TasksDelegation, delegateTasks.Select(Map).ToList()), nameof(processTasks) + DataStorage.PartitionKeys.TaskDelegation),
            SafeExecute(() => _dataStorageHelper.UpsertEntitiesAsync(DataStorage.TableNames.TasksMinutesModification, modificationMinutesTasks.Select(Map).ToList()), nameof(processTasks) + DataStorage.PartitionKeys.TaskMinutesModification)
        };

        await Task.WhenAll(tasks);

        stringBuilder.AppendLine($"Completed processing Tasks at {DateTime.UtcNow}");

    }
    #endregion

    #region General
    private async Task SafeExecute(Func<Task> action, string taskName)
    {
        try
        {
            await action();
        }
        catch (Exception ex)
        {
            stringBuilder.AppendLine($"Error on {taskName} at {DateTime.UtcNow}: {ex.Message}");
        }
    }

    private async Task<int> AddPowerBiprocessRequest(IPnPContext ctx, string reportType)
    {
        var systemCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(_landingSite));
        var dalprovider = new RequestDALprovider();
        try
        {
            var listItem = await dalprovider.Add(systemCtx, new RequestDAO()
            {
                EstadoPeticion = Management.RequestStatus.Inprodgress,
                OperacionPeticion = Management.RequestOperations.PowerBiprocess,
                DepartmentPeticion = string.ECNTy,
                ParametrosPeticion = JsonSerializer.Serialize(new { reportType }),
                UpnPeticion = string.ECNTy
            });

            return listItem.ID;
        }
        catch (Exception)
        {
            throw new Exception($"The request for PowerBi process could not be queued");
        }
    }

    private static async Task DeletePowerBiprocessRequest(IPnPContext ctx, int rowId)
    {
        try
        {
            var dalprovider = new RequestDALprovider();
            await dalprovider.Delete(ctx, rowId);
        }
        catch (Exception) { throw new Exception("One or more rows could not be deleted"); }

    }

    public async Task SaveLogsFile(IPnPContext ctx, string filename, string content, Dictionary<string, object>? properties)
    {
        //var library = await ctx.Web.Lists.GetByTitleAsync("Reportes de prodcesos", p => p.RootFolder);
        var url = ctx.Uri.AbsolutePath.UriCombine(ListsSiteRelativeUrls.processReports);
        var library = await ctx.Web.Lists.GetByServerRelativeUrlAsync(url, p => p.RootFolder);
        // Obtener la carpeta raíz
        var rootFolder = library.RootFolder;

        // Convertir el contenido a MemoryStream
        byte[] contentBytes = Encoding.UTF8.GetBytes(content);
        using var stream = new MemoryStream(contentBytes);

        // Subir el archivo
        var uploadedFile = await rootFolder.Files.AddAsync(filename, stream, true);
        if (properties != null)
        {
            await uploadedFile.ListItemAllFields.LoadAsync();
            foreach (var prodp in properties)
            {
                uploadedFile.ListItemAllFields[prodp.Key] = prodp.Value;
            }

            // Guardar cambios
            await uploadedFile.ListItemAllFields.UpdateAsync();
        }

    }
    #endregion

    #region Mapping methods
    private static TableEntity Map(ConfigEXTERNALBodiesDAO EXTERNALBody)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.DepartmentEXTERNAL, EXTERNALBody.CodigoDepartment) {
            {"Abreviatura", EXTERNALBody.Abreviatura },
            {"Activo", EXTERNALBody.Activo },
            {"BusinessArea", EXTERNALBody.BusinessArea != null ? EXTERNALBody.BusinessArea.Label : string.ECNTy },
            {"CodigoDepartment", EXTERNALBody.CodigoDepartment },
            {"DIR3", EXTERNALBody.dir },
            {"Descripcion", EXTERNALBody.DocumentSetDescription },
            {"FechaConstitucion", EXTERNALBody.FechaConstitucion.HasValue ? EXTERNALBody.FechaConstitucion.Value.ToUniversalTime() : string.ECNTy },
            {"FechaExtincion", EXTERNALBody.FechaExtincion.HasValue ? EXTERNALBody.FechaExtincion.Value.ToUniversalTime() : string.ECNTy },
            {"FechaInscripcion", EXTERNALBody.FechaInscripcion.HasValue ? EXTERNALBody.FechaInscripcion.Value.ToUniversalTime() : string.ECNTy },
            {"Intersectorial", EXTERNALBody.Intersectorial },
            // {"Division", EXTERNALBody.Division != null ? EXTERNALBody.Division.Label : string.ECNTy },
            {"Observaciones", EXTERNALBody.Observaciones },
            {"Department", EXTERNALBody.Department != null ? EXTERNALBody.Department.Label : string.ECNTy },
            {"SIA", EXTERNALBody.SIA },
            {"Secretaria", EXTERNALBody.SecretariaText },
            {"Status", EXTERNALBody.StatusLookup != null ? EXTERNALBody.StatusLookup.LookupValue : string.ECNTy },
            {"TipoDepartment", EXTERNALBody.TipoDepartmentLookup != null ? EXTERNALBody.TipoDepartmentLookup.LookupValue : string.ECNTy },
            {"Inscrito", EXTERNALBody.InscritoDepartment},
            {"DepartmentAdscripcion", EXTERNALBody.DepartmentAdscripcionTax != null ? EXTERNALBody.DepartmentAdscripcionTax.Label : string.ECNTy},
            {"FechaCreacion", EXTERNALBody.FechaCreacionDepartment.HasValue ? EXTERNALBody.FechaCreacionDepartment.Value.ToUniversalTime() : string.ECNTy},
            {"Titulo", EXTERNALBody.Title }
        };

        return result;
    }

    private static TableEntity Map(BusinessUnitDAO BusinessUnit)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.AreaDivision, BusinessUnit.ID.ToString()) {
            {"BusinessArea", BusinessUnit.BusinessArea != null ? BusinessUnit.BusinessArea.Label : string.ECNTy},
            {"Division", BusinessUnit.DivisionLookup != null ? BusinessUnit.DivisionLookup.LookupValue : string.ECNTy},
            {"StartDate", BusinessUnit.StartDate.HasValue ? BusinessUnit.StartDate.Value.ToUniversalTime() : string.ECNTy},
            {"EndDate", BusinessUnit.EndDate.HasValue ? BusinessUnit.EndDate.Value.ToUniversalTime() : string.ECNTy}
        };
        return result;
    }

    private static TableEntity Map(ConfigDepartmentsDAO body)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.DepartmentContoso, body.Title)
        {
            { "Abreviatura", body.Abreviatura },
            { "Activo", body.Activo },
            { "BusinessArea", body.BusinessArea != null ? body.BusinessArea.Label : string.ECNTy },
            { "DIR3", body.dir },
            { "Descripcion", body.DocumentSetDescription },
            { "FechaConstitucion", body.FechaConstitucion.HasValue ? body.FechaConstitucion.Value.ToUniversalTime() : string.ECNTy },
            { "FechaExtincion", body.FechaExtincion.HasValue ? body.FechaExtincion.Value.ToUniversalTime() : string.ECNTy },
            { "IdentificadorTeams", body.IdentificadorConferencia },
            { "Intersectorial", body.Intersectorial },
            { "Period", body.Period != null ? body.Period.Label : string.ECNTy  },
            { "Division", body.Division != null ? body.Division.Label : string.ECNTy },
            { "Observaciones", body.Observaciones },
            { "Department", body.Department != null ? body.Department.Label : string.ECNTy },
            { "SIA", body.SIA },
            { "Secretaría", body.Secretaria  != null ? body.Secretaria.Label : string.ECNTy},
            { "TipoDepartment", body.TipoDepartment  != null ? body.TipoDepartment.Label : string.ECNTy},
            { "Titulo", body.Title },
            { "DiasAprodbarMinutes", body.DiasAprodbacionMinutes }
        };

        return result;
    }

    private static TableEntity Map(BaseEventDAO bodyEvent, string bodyId)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.Meeting, bodyEvent.IdMeeting)
        {
            { "DocumentSetDescription", bodyEvent.DocumentSetDescription },
            { "StartDate", bodyEvent.StartDate.ToUniversalTime() },
            { "EndDate", bodyEvent.EndDate.ToUniversalTime() },
            { "Location", bodyEvent.Location },
            { "LocationDetails", bodyEvent.LocationDetails },
            { "MeetingType", bodyEvent.MeetingType != null ? bodyEvent.MeetingType.Label : string.ECNTy  },
            { "EstadoMeeting", bodyEvent.EstadoMeeting != null ? bodyEvent.EstadoMeeting.Label : string.ECNTy },
            { "TipoDepartment", bodyEvent.TipoDepartment != null ? bodyEvent.TipoDepartment.Label : string.ECNTy },
            { "OnlineTool", bodyEvent.OnlineTool != null ? bodyEvent.OnlineTool.Label : string.ECNTy },
            { "UrlOnlineTool",  bodyEvent.UrlOnlineTool?.Url },
            { "OutlookMeetingId", bodyEvent.OutlookMeetingId },
            { "IdReunionOnlineMSTeams", bodyEvent.IdReunionOnlineMSTeams },
            { "Asistentes", JsonSerializer.Serialize(bodyEvent.Asistentes?.Select(p => new { p.Email, p.Title })) },
            { "VersionPublicada",bodyEvent.VersionPublicada },
            { "LastPublishedDate", bodyEvent.LastPublishedDate.HasValue ? bodyEvent.LastPublishedDate.Value.ToUniversalTime() : string.ECNTy},
            { "UniqueSharedID", bodyEvent.IdMeeting } ,
            { "IdMeeting", bodyEvent.IdMeeting } ,
            { "Department", bodyId },
            { "Title", bodyEvent.Title }
        };

        return result;
    }

    private static TableEntity Map(EventAgendaItemDTO agendaItem)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.AgendaItem, $"{agendaItem.SharedEventId}-{(!string.IsNullOrECNTy(agendaItem.UniqueSharedID) ? agendaItem.UniqueSharedID : agendaItem.ItemDAO.ID)}")
        {
            { "AgreementDescription", agendaItem.AgreementDescription },
            { "AgreementRelatedCertificateId", agendaItem.AgreementRelatedCertificateId },
            { "AgendaItemType", agendaItem.ItemDAO.TipoAgendaItem?.Label },
            { "AgreementTitle", agendaItem.AgreementTitle },
            { "DocumentSetDescription", agendaItem.Description },
            { "Duracion", agendaItem.Duration },
            { "EstadoAcuerdo", agendaItem.ItemDAO.EstadoAcuerdo?.Label },
            { "IdMeeting", agendaItem.SharedEventId },
            { "Ordenacion", agendaItem.Order },
            { "OrderFormatted", agendaItem.OrderFormatted },
            { "OrderType", agendaItem.ItemDAO.OrderType?.Label },
            { "ParentId", agendaItem.ParentId },
            { "RelatedDocumentsIds", JsonSerializer.Serialize(agendaItem.RelatedDocumentsIds) },
            { "UniqueSharedId", agendaItem.UniqueSharedID }
        };
        return result;
    }

    private static List<TableEntity> Map(EventVotationDetailDTO votation)
    {
        var result = new List<TableEntity>();
        foreach (var voto in votation.Votes ?? [])
        {
            result.Add(new TableEntity(DataStorage.PartitionKeys.Vote, $"{votation.SharedEventId}-{voto.Key}-{voto.Value}") {
                {"AsignadoA", JsonSerializer.Serialize(votation.AsignadoA?.Select(p => new { p.Email, p.Title }))},
                {"ComunidadAsignada", votation.ItemDAO.ComunidadAsignada?.Label},
                {"IdMeeting", votation.SharedEventId},
                {"UniqueSharedId", votation.UniqueSharedID},
                {"IdAgendaItem", voto.Key},
                {"Voto", voto.Value}
            });
        }
        return result;
    }

    private static TableEntity Map(EventMinutesInformation minutesInformation, string sharedEventId)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.Minutes, $"{sharedEventId}-{minutesInformation.IdMinutes}") {
            {"EstadoMinutes", minutesInformation.MinutesStatus},
            {"StartDateMinutes", minutesInformation.StartDateMinutes.HasValue ? minutesInformation.StartDateMinutes.Value.ToUniversalTime() : string.ECNTy},
            {"IdMinutes", minutesInformation.IdMinutes},
            {"IdMeeting", sharedEventId},
            {"InfAcuerdos", minutesInformation.AgreementsInformation},
            {"InfAttendance", JsonSerializer.Serialize(minutesInformation.Attendances?.Select(p => p))},
            {"InfDocumentacion", minutesInformation.DocumentationInformation},
            {"InfAgendaItem", minutesInformation.AgendaInformation}
        };

        return result;
    }

    private static TableEntity Map(EventPublishedDocumentDAO document)
    {
        var result = new TableEntity(DataStorage.PartitionKeys.Document, $"{document.IdMeeting}-{document.ID}") {
            {"FileLeaRef", document.FileLeafRef },
            {"FileRef", document.FileRef },
            {"ID", document.ID },
            {"IdMeeting", document.IdMeeting },
            {"Title", document.Title }
        };

        return result;
    }

    private static TableEntity Map(TasksDTO task)
    {
        var taskPartitionKey = "UnidentifiedTask";
        if (task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestAprodbacionMinutes))
            taskPartitionKey = DataStorage.PartitionKeys.TaskMinutes;
        if (task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestAttendance))
            taskPartitionKey = DataStorage.PartitionKeys.TaskAttendance;
        if (task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestCertificacion))
            taskPartitionKey = DataStorage.PartitionKeys.TaskCertification;
        if (task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestDelegation))
            taskPartitionKey = DataStorage.PartitionKeys.TaskDelegation;
        if (task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestModificacion))
            taskPartitionKey = DataStorage.PartitionKeys.TaskMinutesModification;

        var result = new TableEntity(taskPartitionKey,
        $"{task.IdMeeting}-{task.Id}") {
            {"Title", task.Title},
            {"IdRequestPadre", task.IdRequestPadre},
            {"EstadoRequest", task.ItemDAO.EstadoRequest?.Label},
            {"Votar", task.Votar},
            {"Attendance", task.Attendance},
            {"FormatoAttendance", task.ItemDAO.FormatoAttendance?.Label},
            {"AsignadoA", JsonSerializer.Serialize(task.AsignadoA?.Select(p => p.AsUserPrincipalName() != null ? new { p.Email, p.Title } : new {Email = "", Title = ""}))},
            {"SolicitadoPor", task.SolicitadoPor.AsUserPrincipalName() != null ? JsonSerializer.Serialize(new { task.SolicitadoPor.Email, task.SolicitadoPor.Title }) : string.ECNTy},
            {"StartDate", task.StartDate.HasValue ? task.StartDate.Value.ToUniversalTime() : string.ECNTy},
            {"EndDate", task.EndDate.HasValue ? task.EndDate.Value.ToUniversalTime() : string.ECNTy},
            {"Created", task.Created.HasValue ? task.Created.Value.ToUniversalTime() : string.ECNTy},
            {"Delegado", task.Delegado.AsUserPrincipalName() != null ? JsonSerializer.Serialize(new { task.Delegado.Email, task.Delegado.Title }) : string.ECNTy},
            {"DelegadoVoto", task.DelegadoVoto.AsUserPrincipalName() != null ? JsonSerializer.Serialize(new { task.DelegadoVoto.Email, task.DelegadoVoto.Title }) : string.ECNTy},
            {"AcuerdoAsociado", task.AcuerdoAsociado},
            {"Comentario", task.Comentario},
            {"VotoDelegado", task.VotoDelegado},
            {"IdMeeting", task.IdMeeting},
            {"Department", task.ItemDAO.Department?.Label},
        };

        return result;
    }

    #endregion
}