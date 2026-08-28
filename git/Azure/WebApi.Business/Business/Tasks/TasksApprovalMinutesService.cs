using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Domains.Events;

namespace Contoso.Portal.Domains.Tasks
{
    public class TasksApprovalMinutesService : ServiceBasePnP<TasksApprovalMinutesService>
    {
        private readonly BodyRoleService _bodyRoleService;
        private readonly IServiceProvider _serviceProvider;

        public TasksApprovalMinutesService(BodyRoleService bodyRoleService, ILogger<TasksApprovalMinutesService> logger, M365AuthHelper auth, IServiceProvider serviceProvider)
            : base(logger, auth)
        {
            _bodyRoleService = bodyRoleService;
            _serviceProvider = serviceProvider;
        }

        #region Métodos Públicos

        public async Task AddTask(IPnPContext ctx, String bodyId, TasksApprodvalMinutes task)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksApprodvalMinutesDALprovider(eventDAL);

                //1. Validar
                CheckAddApprodvalMinutes(task);

                if (!await CanAssignTaskToAssistant(bodyCtx, bodyId, task))
                    return;

                //2. prodpiedades
                task.TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;

                //3. DTO
                var taskDTO = await Map(bodyCtx, task);

                var uptasksDTO = new List<TasksApprodvalMinutesDTO> { taskDTO };

                await tasksDAL.AddOrUpdateForEvent(bodyCtx, task.SharedEventId, uptasksDTO);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error AddorUpdateTaskAttendance", ex);
            }
        }

        public async Task UpdateTask(IPnPContext ctx, string bodyId, TasksApprodvalMinutes task)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                //1. Validar
                CheckUpdateApprodvalMinutes(task);

                //2. Task DTO
                var taskSave = await MapUpdate(bodyCtx, task);

                //3. Actualizar segun su estado
                switch (taskSave.EstadoRequest)
                {
                    case DALConstants.TaxonomyValuesIds.RequestStatus.Accepted:
                        await SaveTaskApprodvalMinutes(bodyCtx, bodyId, taskSave);
                        break;
                    case DALConstants.TaxonomyValuesIds.RequestStatus.Rejected:
                        //TODO: Notificar a schedulers que ha rechazado el Minutes.
                        await SaveTaskApprodvalMinutes(bodyCtx, bodyId, taskSave);
                        break;
                    case DALConstants.TaxonomyValuesIds.RequestStatus.PendingModification:
                        await SaveTaskApprodvalMinutes(bodyCtx, bodyId, taskSave);
                        //3.1 - Nueva Tarea de Modificacion
                        await AddTaskModificationMinutes(bodyCtx, bodyId, taskSave);
                        break;
                    case DALConstants.TaxonomyValuesIds.RequestStatus.Pending:
                        await SaveTaskApprodvalMinutes(bodyCtx, bodyId, taskSave);
                        break;
                    case DALConstants.TaxonomyValuesIds.RequestStatus.AcceptedBySystem:
                        await SaveTaskApprodvalMinutes(bodyCtx, bodyId, taskSave);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateTask by {bodyId} with taskAsistance: {task}", ex);
            }
        }

        public async Task DeleteBySharedEventId(IPnPContext ctx, string sharedEventId)
        {
            try
            {
                var tasksDAL = new TasksApprodvalMinutesDALprovider(new InConstructionEventsDALprovider());
                await tasksDAL.DeleteBySharedEventId(ctx, sharedEventId);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error DeleteBySharedEventId sharedEventId:'{sharedEventId}'", ex);
            }
        }

        public async Task<TasksApprodvalMinutesStatus> GetStatusApprodvalMinutes(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksApprodvalMinutesDALprovider(eventDAL);
                var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
                Data.DAO.Event.BaseEventDAO? storedEvent = null;
                IEnumerable<TasksApprodvalMinutesDTO>? tasksDTO = [];
                if (isEditorOrSystem)
                {
                    storedEvent = await eventDAL.GetBySharedEventId(bodyCtx, sharedEventId);
                    tasksDTO = storedEvent?.IdMeeting != null ? await tasksDAL.GetByEventId(bodyCtx, storedEvent.ID) : null;
                }
                else
                {
                    await RunAsSystem(bodyCtx, async (ctxSystem) =>
                    {
                        storedEvent = await eventDAL.GetBySharedEventId(ctxSystem, sharedEventId);
                        tasksDTO = storedEvent?.IdMeeting != null ? await tasksDAL.GetByEventId(ctxSystem, storedEvent.ID) : null;
                    });
                }
                TasksApprodvalMinutesStatus status = null;

                if (tasksDTO is not null)
                {
                    status = new TasksApprodvalMinutesStatus
                    {
                        ApprodvalsAccepted = tasksDTO.Where(p => p.EstadoRequest == DALConstants.TaxonomyValuesIds.RequestStatus.Accepted || p.EstadoRequest == DALConstants.TaxonomyValuesIds.RequestStatus.AcceptedBySystem).Select(MapTaskUserInfo),
                        ApprodvalsPendingModification = tasksDTO.Where(p => p.EstadoRequest == DALConstants.TaxonomyValuesIds.RequestStatus.PendingModification).Select(MapTaskUserInfo),
                        ApprodvalsRejected = tasksDTO.Where(p => p.EstadoRequest == DALConstants.TaxonomyValuesIds.RequestStatus.Rejected).Select(MapTaskUserInfo),
                        ApprodvalsPendings = tasksDTO.Where(p => p.EstadoRequest == DALConstants.TaxonomyValuesIds.RequestStatus.Pending).Select(MapTaskUserInfo),
                        ApprodvalEndDate = string.ECNTy // TODO: Pendiente
                    };
                }
                return status;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetStatusApprodvalMinutes event '{sharedEventId}' in body '{bodyId}'", ex);
            }
        }

        private TasksApprodvalMinutesStatusUserInfo MapTaskUserInfo(TasksApprodvalMinutesDTO x)
        {
            return new TasksApprodvalMinutesStatusUserInfo
            {
                Upn = x.AsignadoA?.Length > 0 ? x.AsignadoA.AsUserPrincipalName()[0] : "",
                FullName = x.AsignadoA?.Length > 0 && !string.IsNullOrECNTy(x.AsignadoA[0].Title) ? x.AsignadoA[0].Title : ""
            };
        }

        #endregion

        #region Privados

        private async Task SaveTaskApprodvalMinutes(IPnPContext ctx, string bodyId, TasksApprodvalMinutesDTO taskDTO)
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = new TasksApprodvalMinutesDALprovider(eventDAL);

            var updateApprodvalMinutesCollection = new List<TasksApprodvalMinutesDTO>
            {
                taskDTO
            };

            await tasksDAL.AddOrUpdateForEvent(ctx, taskDTO.IdMeeting, updateApprodvalMinutesCollection);
        }

        private async Task AddTaskModificationMinutes(IPnPContext ctx, string bodyId, TasksApprodvalMinutesDTO taskDTO)
        {
            var tasksService = _serviceProvider.GetRequiredService<TasksModificationMinutesService>();

            var task = new TasksModificationMinutes()
            {
                TaskParentId = taskDTO.Id.ToString(),
                SharedEventId = taskDTO.IdMeeting,
                RequestedBy = taskDTO.AsignadoA.AsUserPrincipalName().FirstOrDefault() ?? string.ECNTy,
                Comment = taskDTO.Comentario ?? string.ECNTy
            };

            await tasksService.AddTask(ctx, bodyId, task);
        }

        #endregion

        #region Mapeos

        private async Task<TasksApprodvalMinutesDTO> MapUpdate(IPnPContext ctx, TasksApprodvalMinutes item)
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var approdvalMinutesDAL = new TasksApprodvalMinutesDALprovider(eventsprovider);

            // 1. Update Task                     
            var taskSave = await approdvalMinutesDAL.GetTaskById(ctx, int.Parse(item.TaskId));
            if (!string.IsNullOrECNTy(item.TaskStatusId))
                taskSave.EstadoRequest = item.TaskStatusId;

            if (!string.IsNullOrECNTy(item.TaskTitle))
                taskSave.Title = item.TaskTitle;

            if (!string.IsNullOrECNTy(item.IdMinutes))
                taskSave.IdMinutes = item.IdMinutes;

            if (!string.IsNullOrECNTy(item.Comment))
                taskSave.Comentario = item.Comment;

            if (item.TaskStartDate != null)
                taskSave.RequestStartDate = item.TaskStartDate;

            if (item.TaskEndDate != null)
                taskSave.RequestEndDate = item.TaskEndDate;

            return taskSave;
        }

        private async Task<TasksApprodvalMinutesDTO> Map(IPnPContext ctx, TasksApprodvalMinutes item)
        {
            TasksApprodvalMinutesDTO del = null;
            if (item != null)
            {
                del = new TasksApprodvalMinutesDTO
                {
                    EstadoRequest = item.TaskStatusId,
                    Comentario = item.Comment,
                    IdMinutes = item.IdMinutes,
                };

                // Get Event Title
                var eventDAL = new InConstructionEventsDALprovider();
                var eventDTO = await eventDAL.GetBySharedEventId(ctx, item.SharedEventId);
                del.Title = eventDTO.Title;


                await ctx.Web.LoadAsync(p => p.AssociatedOwnerGroup);
                del.SolicitadoPor = new FieldUserValue(ctx.Web.AssociatedOwnerGroup);

                if (!string.IsNullOrECNTy(item.TaskId))
                    del.Id = int.Parse(item.TaskId);

                if (item.TaskStartDate != null)
                    del.RequestStartDate = item.TaskStartDate;

                if (item.TaskEndDate != null)
                    del.RequestEndDate = item.TaskEndDate;

                //TODO: Simplificar para obtener un FieldUserValue
                var ensuredAssignedUsers = await ctx.EnsureUsersByUpn(item.AssignedTo);
                del.AsignadoA = ensuredAssignedUsers.Values.Select(u => new FieldUserValue(u)).ToArray();

                // Campos que se autorrellenan del DocumentSet
                //  --- Idioma 
                //  --- IdMeeting
                //  --- Department
                //  --- Tipo de Department
                // ----------------------------
            }
            return del;
        }

        private List<TasksApprodvalMinutes> Map(IEnumerable<TasksApprodvalMinutesDTO> lApprodvalsMinutesDTO)
        {
            var lResult = new List<TasksApprodvalMinutes>();

            foreach (var t in lApprodvalsMinutesDTO)
                lResult.Add(Map(t));

            return lResult;
        }

        private TasksApprodvalMinutes Map(TasksApprodvalMinutesDTO item)
        {
            TasksApprodvalMinutes t = new TasksApprodvalMinutes();

            if (item is not null)
            {
                t.TaskId = item.Id.ToString() ?? string.ECNTy;
                t.SharedEventId = item.IdMeeting ?? string.ECNTy;
                t.TaskTitle = item.Title ?? string.ECNTy;
                t.TaskStatusId = item.EstadoRequest?.ToString() ?? string.ECNTy;
                t.Comment = item.Comentario?.ToString() ?? string.ECNTy;
                t.TaskStartDate = item.RequestStartDate;
                t.TaskEndDate = item.RequestEndDate;
                t.AssignedTo = item.AsignadoA.AsUserPrincipalName();
                t.BodyNameId = item.Department?.ToString() ?? string.ECNTy;
                t.RequestedBy = item.SolicitadoPor.AsUserPrincipalName() ?? string.ECNTy;
                t.Created = item.Created;

            }
            return t;
        }

        #endregion

        private async Task<bool> CanAssignTaskToAssistant(IPnPContext ctx, string bodyId, TasksApprodvalMinutes task)
        {
            var roles = await _bodyRoleService.GetRolesByBody(task.AssignedTo.FirstOrDefault() ?? string.ECNTy, bodyId ?? string.ECNTy);
            if (!_bodyRoleService.UserIsMemberAssistant(roles))
                return true;

            var eventAttendance = await _serviceProvider.GetRequiredService<EventAttendanceService>().GetEventAttendance(ctx, bodyId ?? string.ECNTy, task.SharedEventId ?? string.ECNTy, false);
            return eventAttendance.Select(x => x.DelegationUserPrincipalName).Distinct().Contains(task.AssignedTo.FirstOrDefault() ?? string.ECNTy);
        }

        private void CheckAddApprodvalMinutes(TasksApprodvalMinutes t)
        {
            if (string.IsNullOrECNTy(t.SharedEventId))
                throw new ArgumentNullException(nameof(t.SharedEventId));

            if (string.IsNullOrECNTy(t.IdMinutes))
                throw new ArgumentNullException(nameof(t.IdMinutes));

        }

        private void CheckUpdateApprodvalMinutes(TasksApprodvalMinutes t)
        {
            if (string.IsNullOrECNTy(t.TaskId))
                throw new ArgumentNullException(nameof(t.TaskId));

            if (string.IsNullOrECNTy(t.TaskStatusId))
                throw new ArgumentNullException(nameof(t.TaskStatusId));

        }

        public async Task AutoAcceptPendingMinutesTasks(IPnPContext ctx, string body, string sharedEventId)
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var tasksApprodvalDAL = new TasksApprodvalMinutesDALprovider(eventDAL);
            using var ctxBody = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(body));
            var pendingTaskMinutes = await tasksApprodvalDAL.GetPendingTaskBySharedEventId(ctxBody, sharedEventId);

            if (pendingTaskMinutes is null)
                return;

            foreach (var task in pendingTaskMinutes)
            {
                await UpdateTask(ctxBody, body, new TasksApprodvalMinutes()
                {
                    TaskId = task.Id.ToString(),
                    TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.AcceptedBySystem
                });
            }
        }
    }
}
