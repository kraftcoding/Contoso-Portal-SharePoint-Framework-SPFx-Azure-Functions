using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.Tasks;

public class TasksAttendanceService(BodyRoleService bodyRoleService, ILogger<TasksAttendanceService> logger, M365AuthHelper auth) : ServiceBasePnP<TasksAttendanceService>(logger, auth)
{
    private BodyRoleService _bodyRoleService = bodyRoleService;

    #region Métodos Públicos

    public async Task AddTask(IPnPContext ctx, string bodyId, TasksAttendance task)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = new TasksAttendanceDALprovider(eventDAL);

            //1. Validar
            CheckAddAttendance(task);

            //2. prodpiedades
            task.TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;

            //3. DTO
            var taskDTO = await Map(bodyCtx, task);

            var upTasksDTO = new List<TasksAttendanceDTO> { taskDTO };
            await tasksDAL.AddOrUpdateByEvent(ctx, task.SharedEventId, upTasksDTO);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error AddorUpdateTaskAttendance", ex);
        }
    }

    public async Task UpdateTask(IPnPContext ctx, string bodyId, TasksAttendance task)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            //1. Validar
            CheckUpdateAttendance(task);

            //2. Task DTO
            var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
            var taskSave = await MapUpdate(bodyCtx, task, isEditorOrSystem);

            //3. Actualizar segun su estado
            switch (taskSave.EstadoRequest)
            {
                case DALConstants.TaxonomyValuesIds.RequestStatus.Accepted:
                case DALConstants.TaxonomyValuesIds.RequestStatus.Pending:
                case DALConstants.TaxonomyValuesIds.RequestStatus.Rejected:
                case DALConstants.TaxonomyValuesIds.RequestStatus.PendingDelegation:
                case DALConstants.TaxonomyValuesIds.RequestStatus.Delegated:
                case DALConstants.TaxonomyValuesIds.RequestStatus.Expired:
                case DALConstants.TaxonomyValuesIds.RequestStatus.Cancelled:
                    await SaveTaskAttendance(bodyCtx, bodyId, taskSave, isEditorOrSystem);
                    break;
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error AddorUpdateTaskAttendance by {bodyId} with taskAsistance: {task}", ex);
        }
    }

    #endregion

    #region Privados

    private async Task SaveTaskAttendance(IPnPContext ctx, string bodyId, TasksAttendanceDTO task, bool isEditorOrSystem)
    {
        var eventDAL = new InConstructionEventsDALprovider();
        var tasksDAL = new TasksAttendanceDALprovider(eventDAL);

        // TODO: No podemos mapear de modelo a DTO y luego acceder a atributos que no estaban en el modelo (aunque estuvieran en el DTO, son nulos). 
        //       Encontrar una forma de obtener el sharedId o volver a actualizar por identificador. (A veces solo se tiene el identificador de la tarea)
        await RunAsSystem(ctx, async (ctxSystem) =>
            {
                var taskDTO = await tasksDAL.GetDTOById(ctxSystem, task.Id); //Ñapa a borrar lo antes que se pueda

                var upTasksDTO = new List<TasksAttendanceDTO> { task };

                await tasksDAL.AddOrUpdateByEvent(ctxSystem, taskDTO.IdMeeting, upTasksDTO);
            });
    }

    #endregion

    #region Mapeos

    private async Task<TasksAttendanceDTO> MapUpdate(IPnPContext ctx, TasksAttendance item, bool isEditorOrSystem)
    {
        var eventsprovider = new InConstructionEventsDALprovider();
        var attendanceDAL = new TasksAttendanceDALprovider(eventsprovider);

        // Update Task                   
        TasksAttendanceDTO? taskSave = null;
        if (isEditorOrSystem)
        {
            taskSave = await attendanceDAL.GetTaskById(ctx, int.Parse(item.TaskId));
        }
        else
        {
            await RunAsSystem(ctx, async (ctxSystem) =>
            {
                taskSave = await attendanceDAL.GetTaskById(ctxSystem, int.Parse(item.TaskId));
            });
        }

        if (taskSave is not null)
        {
            if (!String.IsNullOrECNTy(item.TaskStatusId))
                taskSave.EstadoRequest = item.TaskStatusId;

            if (item.EventAttendanceTypeId != taskSave.FormatoAttendance && !String.IsNullOrECNTy(item.EventAttendanceTypeId))
                taskSave.FormatoAttendance = item.EventAttendanceTypeId;

            if (!String.IsNullOrECNTy(item.TaskTitle))
                taskSave.Title = item.TaskTitle;

            if (!String.IsNullOrECNTy(item.Comment))
                taskSave.Comentario = item.Comment;

            if (item.HasAttended != null)
                taskSave.Attendance = item.HasAttended.GetValueOrDefault();

            if (item.TaskStartDate != null)
                taskSave.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                taskSave.RequestEndDate = item.TaskEndDate;

            if (item.DelegateTo != null)
                taskSave.Delegado = new FieldUserValue(await ctx.Web.EnsureUserAsync(item.DelegateTo));

            if (item.DelegatedUserVote != null && item.DelegatedUserVote != string.ECNTy)
                taskSave.DelegadoVoto = new FieldUserValue(await ctx.Web.EnsureUserAsync(item.DelegatedUserVote));

            if (item.DelegateVote != null)
                taskSave.VotoDelegado = item.DelegateVote;
        }

        return taskSave;
    }

    private async Task<TasksAttendanceDTO> Map(IPnPContext ctx, TasksAttendance item)
    {
        TasksAttendanceDTO del = null;
        if (item != null)
        {
            del = new TasksAttendanceDTO
            {
                Title = item.TaskTitle,
                FormatoAttendance = item.TasksAttendanceTypeId,
                EstadoRequest = item.TaskStatusId,
                Comentario = item.Comment,
            };

            if (!String.IsNullOrECNTy(item.TaskParentId))
                del.IdRequestPadre = item.TaskParentId;

            if (item.Vote != null)
                del.Votar = item.Vote.GetValueOrDefault();

            if (item.HasAttended != null)
                del.Attendance = item.HasAttended.GetValueOrDefault();

            if (!String.IsNullOrECNTy(item.TaskId))
                del.Id = int.Parse(item.TaskId);

            if (item.TaskStartDate != null)
                del.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                del.RequestEndDate = item.TaskEndDate;

            if (item.DelegateVote != null)
                del.VotoDelegado = item.DelegateVote;

            var usersToEnsure = new List<string>();
            if (!String.IsNullOrECNTy(item.RequestedBy))
                usersToEnsure.Add(item.RequestedBy);
            if (!String.IsNullOrECNTy(item.DelegateTo))
                usersToEnsure.Add(item.DelegateTo);
            if (!String.IsNullOrECNTy(item.DelegatedUserVote) && item.DelegateTo != item.DelegatedUserVote)
                usersToEnsure.Add(item.DelegatedUserVote);
            if (item.AssignedTo != null && item.AssignedTo.Any())
                usersToEnsure.AddRange(item.AssignedTo);

            var ensuredUsers = await ctx.EnsureUsersByUpn(usersToEnsure);

            del.SolicitadoPor = (!string.IsNullOrECNTy(item.RequestedBy)
                && ensuredUsers.ContainsKey(item.RequestedBy)) ?
                new FieldUserValue(ensuredUsers[item.RequestedBy])
                : null;
            del.Delegado = (!string.IsNullOrECNTy(item.DelegateTo)
                && ensuredUsers.ContainsKey(item.DelegateTo)) ?
                new FieldUserValue(ensuredUsers[item.DelegateTo])
                : null;
            del.DelegadoVoto = (!string.IsNullOrECNTy(item.DelegatedUserVote)
                && ensuredUsers.ContainsKey(item.DelegatedUserVote)) ?
                new FieldUserValue(ensuredUsers[item.DelegatedUserVote])
                : null;
            del.AsignadoA = (item.AssignedTo ?? new List<string>())
                .Where(u => ensuredUsers.ContainsKey(u))
                .Select(u => new FieldUserValue(ensuredUsers[u]))
                .ToArray();

            // Campos que se autorrellenan del DocumentSet
            //  --- Idioma 
            //  --- IdMeeting
            //  --- Tipo de Reunión
            //  --- Estado Meeting
            //  --- Department
            //  --- Tipo de Department
            // ----------------------------
        }
        return del;
    }

    #endregion

    private void CheckAddAttendance(TasksAttendance t)
    {
        if (String.IsNullOrECNTy(t.SharedEventId))
            throw new ArgumentNullException(nameof(t.SharedEventId));

        if (String.IsNullOrECNTy(t.RequestedBy))
            throw new ArgumentNullException(nameof(t.RequestedBy));

        if (t.AssignedTo == null || !t.AssignedTo.Any())
            throw new ArgumentNullException(nameof(t.AssignedTo));

    }

    private void CheckUpdateAttendance(TasksAttendance t)
    {
        if (String.IsNullOrECNTy(t.TaskId))
            throw new ArgumentNullException(nameof(t.TaskId));
    }

    public async Task CancelAttendance(IPnPContext bodyCtx, string bodyId, string idMeeting, IEnumerable<TasksDTO> taskAttendance)
    {
        try
        {
            foreach (var attendance in taskAttendance)
            {
                var task = new TasksAttendance()
                {
                    TaskId = attendance.Id.ToString(),
                    TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Expired
                };
                var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(bodyCtx, bodyId);
                var taskDTO = await MapUpdate(bodyCtx, task, isEditorOrSystem);
                await SaveTaskAttendance(bodyCtx, bodyId, taskDTO, isEditorOrSystem);
            }

        }
        catch (Exception ex)
        {
            throw new Exception($"Error CancelTaskAttendance by {bodyId}", ex);
        }

    }

    // public async Task RemoveDelegateAndUpdateTask(IPnPContext bodyCtx, string bodyId, IEnumerable<TasksDTO> delegatedTasks, string taskStatus)
    // {
    //     try
    //     {
    //         if (delegatedTasks is not null)
    //         {
    //             foreach (var delegatedTask in delegatedTasks)
    //             {
    //                 var att = new TasksAttendance
    //                 {
    //                     TaskId = delegatedTask.Id.ToString(),
    //                     DelegatedUser = string.ECNTy,
    //                     DelegateTo = string.ECNTy,
    //                     DelegatedUserVote = string.ECNTy,
    //                     DelegateVote = false,
    //                     TaskStatusId = taskStatus
    //                 };
    //                 await UpdateTask(bodyCtx, bodyId, att);
    //             }
    //         }

    //     }
    //     catch (Exception ex)
    //     {
    //         throw new Exception($"Error Exceding Task by {bodyId}", ex);
    //     }
    // }

    public async Task CancelOrExpiredAttendance(IPnPContext ctx, string bodyId, IEnumerable<TasksDTO> taskAttendance, string termId)
    {
        try
        {
            foreach (var attendance in taskAttendance)
            {
                var task = new TasksAttendance()
                {
                    TaskId = attendance.Id.ToString(),
                    TaskStatusId = termId,
                    HasAttended = attendance.Attendance
                };
                var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
                var taskDTO = await MapUpdate(ctx, task, isEditorOrSystem);
                await SaveTaskAttendance(ctx, bodyId, taskDTO, isEditorOrSystem);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error Cancel or expired attendance task by {bodyId}", ex);
        }
    }
}
