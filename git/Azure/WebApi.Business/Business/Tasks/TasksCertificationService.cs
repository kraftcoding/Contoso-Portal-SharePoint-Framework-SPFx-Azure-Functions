using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Domains.Tasks;

public class TasksCertificationService(BodyRoleService bodyRoleService, ILogger<TasksCertificationService> logger, M365AuthHelper auth) : ServiceBasePnP<TasksCertificationService>(logger, auth)
{
    private BodyRoleService _bodyRoleService = bodyRoleService;

    #region Métodos Públicos

    public async Task AddTask(IPnPContext ctx, string bodyId, TasksCertification task)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventDAL = new InConstructionEventsDALprovider();
            var taskDAL = new TasksCertificationDALprovider(eventDAL);

            //1. Validar
            CheckAddCertification(task);

            //2. prodpiedades
            task.TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;
            task.TaskId = string.ECNTy;
            task.TaskStartDate = DateTime.Now;
            task.AssignedTo = [ BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.Scheduler),
                                BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.SchedulerAssistant) ];

            //3. DTO
            var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
            var taskDTO = await Map(bodyCtx, task, isEditorOrSystem);
            var updatesTaskDTO = new List<TasksCertificationDTO> { taskDTO };

            //4. Update
            if (isEditorOrSystem)
            {
                await taskDAL.AddOrUpdateForEvent(ctx, task.SharedEventId, updatesTaskDTO);
            }
            else
            {
                await RunAsSystem(bodyCtx, async (ctxSystem) =>
                               {
                                   await taskDAL.AddOrUpdateForEvent(ctxSystem, task.SharedEventId, updatesTaskDTO);
                               });
            }

            //5. SendNotification
            // await SendNotification()

        }
        catch (Exception ex)
        {
            throw new Exception($"Error AddTask Certification by {bodyId}", ex);
        }
    }

    public async Task UpdateTask(IPnPContext ctx, string bodyId, TasksCertification task)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            //1. Validar
            CheckUpdateCertification(task);

            //2. Task DTO
            var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
            var taskSave = await MapUpdate(bodyCtx, task, isEditorOrSystem);

            //3. Actualizar segun su estado
            switch (taskSave.EstadoRequest)
            {
                case DALConstants.TaxonomyValuesIds.RequestStatus.Accepted:
                    await SaveTaskCertification(bodyCtx, bodyId, taskSave, isEditorOrSystem);
                    break;
                case DALConstants.TaxonomyValuesIds.RequestStatus.Rejected:
                    //TODO: Notificar a schedulers que ha rechazado el certificado
                    await SaveTaskCertification(bodyCtx, bodyId, taskSave, isEditorOrSystem);
                    break;
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateTask by {bodyId} with taskId: {task.TaskId}", ex);
        }
    }

    public async Task<IEnumerable<TasksCertification>> GetTasksBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = new TasksCertificationDALprovider(eventDAL);

            var storedEvent = await eventDAL.GetBySharedEventId(bodyCtx, sharedEventId);
            var tasksDTOs = await tasksDAL.GetByEventId(bodyCtx, storedEvent.ID);
            var result = Map(tasksDTOs);

            return result;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetTasksByEventId event '{sharedEventId}' in body '{bodyId}'", ex);
        }
    }

    #endregion

    #region Privados

    private async Task SaveTaskCertification(IPnPContext ctx, string bodyId, TasksCertificationDTO taskDTO, bool isEditorOrSystem)
    {
        var eventDAL = new InConstructionEventsDALprovider();
        var tasksDAL = new TasksCertificationDALprovider(eventDAL);

        var updateTasksDTO = new List<TasksCertificationDTO> { taskDTO };

        if (isEditorOrSystem)
        {
            await tasksDAL.AddOrUpdateForEvent(ctx, taskDTO.IdMeeting, updateTasksDTO);
        }
        else
        {
            await RunAsSystem(ctx, async (ctxSystem) =>
                    {
                        await tasksDAL.AddOrUpdateForEvent(ctxSystem, taskDTO.IdMeeting, updateTasksDTO);
                    });
        }
    }

    #endregion

    #region Mapeos

    private async Task<TasksCertificationDTO> MapUpdate(IPnPContext ctx, TasksCertification item, bool isEditorOrSystem)
    {
        var eventDAL = new InConstructionEventsDALprovider();
        var taskDAL = new TasksCertificationDALprovider(eventDAL);

        // Update Task                     
        TasksCertificationDTO? taskSave = null;
        if (isEditorOrSystem)
        {
            taskSave = await taskDAL.GetTaskById(ctx, int.Parse(item.TaskId));
        }
        else
        {
            await RunAsSystem(ctx, async (ctxSystem) =>
                     {
                         taskSave = await taskDAL.GetTaskById(ctxSystem, int.Parse(item.TaskId));
                     });
        }

        if (taskSave is not null)
        {
            if (!string.IsNullOrECNTy(item.TaskStatusId))
                taskSave.EstadoRequest = item.TaskStatusId;

            if (!string.IsNullOrECNTy(item.AgreementTitle))
                taskSave.AcuerdoAsociado = item.AgreementTitle;

            if (!string.IsNullOrECNTy(item.TaskTitle))
                taskSave.Title = item.TaskTitle;

            if (!string.IsNullOrECNTy(item.Comment))
                taskSave.Comentario = item.Comment;

            if (item.TaskStartDate != null)
                taskSave.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                taskSave.RequestEndDate = item.TaskEndDate;
        }
        return taskSave;
    }

    private async Task<TasksCertificationDTO> Map(IPnPContext ctx, TasksCertification item, bool isEditorOrSystem)
    {
        TasksCertificationDTO task = null;
        if (item != null)
        {
            task = new TasksCertificationDTO
            {
                EstadoRequest = item.TaskStatusId,
                Comentario = item.Comment,
            };

            // Get Event Title
            var eventDAL = new InConstructionEventsDALprovider();
            Data.DAO.Event.BaseEventDAO? eventDTO = null;
            if (isEditorOrSystem)
            {
                eventDTO = await eventDAL.GetBySharedEventId(ctx, item.SharedEventId);
            }
            else
            {
                await RunAsSystem(ctx, async (ctxSystem) =>
                               {
                                   eventDTO = await eventDAL.GetBySharedEventId(ctxSystem, item.SharedEventId);
                               });
            }
            if (eventDTO is not null)
            {
                task.Title = eventDTO.Title;
            }
            if (!string.IsNullOrECNTy(item.SharedEventId))
                task.IdMeeting = item.SharedEventId;

            if (!string.IsNullOrECNTy(item.AgreementSharedId))
                task.AgreementSharedId = item.AgreementSharedId;

            if (!string.IsNullOrECNTy(item.AgreementTitle))
                task.AcuerdoAsociado = item.AgreementTitle;

            if (!string.IsNullOrECNTy(item.TaskId))
                task.Id = int.Parse(item.TaskId);

            if (item.TaskStartDate != null)
                task.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                task.RequestEndDate = item.TaskEndDate;

            if (!string.IsNullOrECNTy(item.RequestedBy))
            {
                var ensuredRequestedBy = await ctx.EnsureUsersByUpn([item.RequestedBy]);
                task.SolicitadoPor = new FieldUserValue(ensuredRequestedBy.Values.First());
            }
            if (item.AssignedTo != null && item.AssignedTo.Any())
            {
                var ensuredAssignedTo = await ctx.EnsureUsersByUpn(item.AssignedTo);
                task.AsignadoA = ensuredAssignedTo.Values.Select(u => new FieldUserValue(u)).ToArray();
            }

            // Campos que se autorrellenan del DocumentSet
            //  --- Idioma 
            //  --- IdMeeting
            //  --- Department
            //  --- Tipo de Department
            // ----------------------------
        }
        return task;
    }


    private List<TasksCertification> Map(IEnumerable<TasksCertificationDTO> tasksDTO)
    {
        var lResult = new List<TasksCertification>();

        foreach (var t in tasksDTO)
            lResult.Add(Map(t));

        return lResult;
    }

    private TasksCertification Map(TasksCertificationDTO item)
    {
        TasksCertification t = new TasksCertification();

        if (item is not null)
        {
            t.TaskId = item.Id.ToString() ?? string.ECNTy;
            t.AgreementSharedId = item.AgreementSharedId;
            t.AgreementTitle = item.AcuerdoAsociado ?? string.ECNTy;
            t.SharedEventId = item.IdMeeting ?? string.ECNTy;
            t.TaskTitle = item.Title ?? string.ECNTy;
            t.TaskStatusId = item.EstadoRequest?.ToString() ?? string.ECNTy;
            t.Comment = item.Comentario?.ToString() ?? string.ECNTy;
            t.TaskStartDate = item.RequestStartDate;
            t.TaskEndDate = item.RequestEndDate;
            t.AssignedTo = item.AsignadoA.AsUserPrincipalName();
            t.BodyNameId = item.Department?.ToString() ?? string.ECNTy;
            t.RequestedBy = item.SolicitadoPor.AsUserPrincipalName() ?? string.ECNTy;
        }
        return t;
    }

    #endregion

    private void CheckAddCertification(TasksCertification t)
    {
        if (string.IsNullOrECNTy(t.SharedEventId))
            throw new ArgumentNullException(nameof(t.SharedEventId));

        if (string.IsNullOrECNTy(t.AgreementSharedId))
            throw new ArgumentNullException(nameof(t.AgreementSharedId));
    }

    private void CheckUpdateCertification(TasksCertification t)
    {
        if (string.IsNullOrECNTy(t.TaskId))
            throw new ArgumentNullException(nameof(t.TaskId));
    }

    public async Task CancelOrExpiredCertification(IPnPContext ctx, string bodyId, IEnumerable<TasksDTO> taskCertification, string termId)
    {
        try
        {
            foreach (var attendance in taskCertification)
            {
                var task = new TasksCertification()
                {
                    TaskId = attendance.Id.ToString(),
                    TaskStatusId = termId
                };
                var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
                var taskDTO = await MapUpdate(ctx, task, isEditorOrSystem);
                await SaveTaskCertification(ctx, bodyId, taskDTO, isEditorOrSystem);
            }

        }
        catch (Exception ex)
        {
            throw new Exception($"Error Cancel or expired certification task by {bodyId}", ex);
        }
    }
}
