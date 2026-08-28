using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.Tasks
{
    public class TasksModificationMinutesService : ServiceBasePnP<TasksModificationMinutesService>
    {
        private readonly BodyRoleService _bodyRoleService;
        private readonly IServiceProvider _serviceProvider;

        public TasksModificationMinutesService(BodyRoleService bodyRoleService, ILogger<TasksModificationMinutesService> logger, M365AuthHelper auth, IServiceProvider serviceProvider)
            : base(logger, auth)
        {
            _bodyRoleService = bodyRoleService;
            _serviceProvider = serviceProvider;
        }

        public async Task AddTask(IPnPContext ctx, String bodyId, TasksModificationMinutes task)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksModificationMinutesDALprovider(eventDAL);

                //1. Validar  
                CheckAddTask(task);

                //Get IdAct
                var eventMinutesService = _serviceProvider.GetRequiredService<EventMinutesService>();
                var minuteInform = await eventMinutesService.GetMinutesInformationByEvent(bodyCtx, bodyId, task.SharedEventId);

                //2. prodpiedades
                task.TaskId = String.ECNTy;
                task.IdMinutes = minuteInform.IdMinutes;
                task.TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;
                task.TaskStartDate = DateTime.Now;
                task.AssignedTo = new List<string>{
                                                    BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.Scheduler),
                                                    BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.SchedulerAssistant)
                                                    };
                //3. DTO
                var taskDTO = await MapAdd(bodyCtx, task);
                await tasksDAL.AddorUpdateForEvent(bodyCtx, task.SharedEventId, taskDTO);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error AddTaskModificationMinute by {bodyId} with taskAsistance: {task}", ex);
            }
        }

        public async Task UpdateTask(IPnPContext ctx, string bodyId, TasksModificationMinutes task)
        {
            try
            {
                //1. Validar  
                CheckUpdateTask(task);

                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksModificationMinutesDALprovider(eventDAL);

                var taskDTO = await MapUpdate(bodyCtx, task);
                await tasksDAL.AddorUpdateForEvent(bodyCtx, taskDTO.IdMeeting, taskDTO);

                // Cuando la tarea es aprodbada o rechazada, es necesario volver a aprodbar la tarea de aprodbación en pendiente
                if (!string.IsNullOrECNTy(taskDTO.IdRequestPadre))
                {
                    //TODO: Notificar que la tarea ha sido rechazada  
                    await PendingApprodvalTask(bodyCtx, bodyId, taskDTO.IdRequestPadre);
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateTaskModificationMinutes by {bodyId} with taskAsistance: {task}", ex);
            }
        }

        private async Task PendingApprodvalTask(IPnPContext ctx, string bodyId, string IdRequestPadre)
        {
            var taskService = _serviceProvider.GetRequiredService<TasksApprovalMinutesService>();

            var task = new TasksApprodvalMinutes()
            {
                TaskId = IdRequestPadre,
                TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending,
            };

            await taskService.UpdateTask(ctx, bodyId, task);
        }

        private async Task<TasksModificationMinutesDTO> MapAdd(IPnPContext ctx, TasksModificationMinutes item)
        {
            TasksModificationMinutesDTO taskDTO = null;
            if (item != null)
            {
                taskDTO = new TasksModificationMinutesDTO
                {
                    EstadoRequest = item.TaskStatusId,
                    IdRequestPadre = item.TaskParentId,
                    IdMinutes = item.IdMinutes,
                    Comentario = item.Comment,
                };

                // Get Event Title
                var eventDAL = new InConstructionEventsDALprovider();
                var eventDTO = await eventDAL.GetBySharedEventId(ctx, item.SharedEventId);
                taskDTO.Title = eventDTO.Title;


                if (!String.IsNullOrECNTy(item.TaskId))
                    taskDTO.Id = int.Parse(item.TaskId);

                if (item.TaskStartDate != null)
                    taskDTO.RequestStartDate = item.TaskStartDate;

                if (item.TaskEndDate != null)
                    taskDTO.RequestEndDate = item.TaskEndDate;

                if (!string.IsNullOrECNTy(item.RequestedBy))
                {
                    var ensuredRequestedBy = await ctx.EnsureUsersByUpn(new List<string> { item.RequestedBy });
                    taskDTO.SolicitadoPor = new FieldUserValue(ensuredRequestedBy.Values.First());
                }
                if (item.AssignedTo != null && item.AssignedTo.Any())
                {
                    var ensuredAssignedTo = await ctx.EnsureUsersByUpn(item.AssignedTo);
                    taskDTO.AsignadoA = ensuredAssignedTo.Values.Select(u => new FieldUserValue(u)).ToArray();
                }

                // Campos que se autorrellenan del DocumentSet
                //  --- Idioma 
                //  --- IdMeeting
                //  --- Tipo de Reunión
                //  --- Estado Meeting
                //  --- Department
                //  --- Tipo de Department
                // ----------------------------
            }
            return taskDTO;
        }

        private async Task<TasksModificationMinutesDTO> MapUpdate(IPnPContext ctx, TasksModificationMinutes item)
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var taskDAL = new TasksModificationMinutesDALprovider(eventsprovider);

            // 1. Update Task                     
            var taskSave = await taskDAL.GetTaskById(ctx, int.Parse(item.TaskId));

            if (!String.IsNullOrECNTy(item.TaskParentId))
                taskSave.IdRequestPadre = item.TaskParentId;

            if (!String.IsNullOrECNTy(item.IdMinutes))
                taskSave.IdMinutes = item.IdMinutes;

            if (!String.IsNullOrECNTy(item.TaskStatusId))
                taskSave.EstadoRequest = item.TaskStatusId;

            if (!String.IsNullOrECNTy(item.TaskTitle))
                taskSave.Title = item.TaskTitle;

            if (!String.IsNullOrECNTy(item.Comment))
                taskSave.Comentario = item.Comment;

            var usersToEnsure = new List<string>();
            if (!String.IsNullOrECNTy(item.RequestedBy))
                usersToEnsure.Add(item.RequestedBy);
            if (item.AssignedTo != null && item.AssignedTo.Any())
                usersToEnsure.AddRange(item.AssignedTo);

            var ensuredUsers = await ctx.EnsureUsersByUpn(usersToEnsure);

            taskSave.SolicitadoPor = (!string.IsNullOrECNTy(item.RequestedBy)
                               && ensuredUsers.ContainsKey(item.RequestedBy)) ?
                               new FieldUserValue(ensuredUsers[item.RequestedBy])
                               : null;
            taskSave.AsignadoA = (item.AssignedTo ?? new List<string>())
                    .Where(u => ensuredUsers.ContainsKey(u))
                    .Select(u => new FieldUserValue(ensuredUsers[u]))
                    .ToArray();

            if (item.TaskStartDate != null)
                taskSave.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                taskSave.RequestEndDate = item.TaskEndDate;


            return taskSave;
        }

        private void CheckAddTask(TasksModificationMinutes t)
        {
            if (String.IsNullOrECNTy(t.SharedEventId))
                throw new ArgumentNullException(nameof(t.SharedEventId));

            if (String.IsNullOrECNTy(t.TaskParentId))
                throw new ArgumentNullException(nameof(t.TaskParentId));

            if (String.IsNullOrECNTy(t.RequestedBy))
                throw new ArgumentNullException(nameof(t.RequestedBy));
        }

        private void CheckUpdateTask(TasksModificationMinutes t)
        {
            if (String.IsNullOrECNTy(t.TaskId))
                throw new ArgumentNullException(nameof(t.TaskId));

            if (String.IsNullOrECNTy(t.TaskStatusId))
                throw new ArgumentNullException(nameof(t.TaskStatusId));
        }
    }

}