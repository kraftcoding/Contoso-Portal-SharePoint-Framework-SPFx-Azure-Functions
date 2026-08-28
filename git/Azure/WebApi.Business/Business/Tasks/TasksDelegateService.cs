using System.IO.Comtestssion;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DAO.Tasks;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.Security;
using PnP.Core.Model.SharePoint;
using PnP.Core.QueryModel;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Domains.Tasks
{
    public class TasksDelegateService : ServiceBasePnP<TasksDelegateService>
    {
        private readonly BodyRoleService _bodyRoleService;

        private readonly IServiceprovider _serviceprovider;

        private readonly NotificationsService _notificationsService;

        private readonly TasksAttendanceService _tasksAttendanceService;

        private readonly string _defaultLocale;

        public TasksDelegateService(BodyRoleService bodyRoleService, ILogger<TasksDelegateService> logger, M365AuthHelper auth, IServiceprovider serviceprovider, NotificationsService notificationsService, TasksAttendanceService tasksAttendanceService, string defaultLocale)
            : base(logger, auth)
        {
            _notificationsService = notificationsService;
            _bodyRoleService = bodyRoleService;
            _serviceprovider = serviceprovider;
            _tasksAttendanceService = tasksAttendanceService;
            _defaultLocale = defaultLocale;
        }

        public async Task AddTask(IPnPContext ctx, String bodyId, TasksDelegate task)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksDelegateDALprovider(eventDAL);

                //1. Validar  
                CheckAddTaskDelegate(task);

                //2. prodpiedades
                task.TaskId = String.ECNTy;
                task.TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;
                task.TaskStartDate = DateTime.Now;
                task.AssignedTo = new List<string>{
                                                    BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.Scheduler),
                                                    BodyPatternUtilities.RoleToGroupName(bodyId,BodyRole.SchedulerAssistant)
                                                    };
                //3. DTO
                var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, bodyId);
                var taskDTO = await MapAddDelegate(bodyCtx, task);
                if (isEditorOrSystem)
                {
                    await tasksDAL.AddorUpdateForEvent(bodyCtx, task.SharedEventId, taskDTO);
                    await UpdateAttendanceTask(bodyCtx, bodyId, task);
                }
                else
                {
                    await RunAsSystem(bodyCtx, async (ctxSystem) =>
                                  {
                                      await tasksDAL.AddorUpdateForEvent(ctxSystem, task.SharedEventId, taskDTO);
                                      await UpdateAttendanceTask(ctxSystem, bodyId, task);
                                  });
                }

                //4. Update AttendanceTask -> PendingDelegattion


                //5. Notifificacion

            }
            catch (Exception ex)
            {
                throw new Exception($"Error AddTaskDelegate by {bodyId} with taskAsistance: {task}", ex);
            }
        }

        public async Task UpdateTask(IPnPContext ctx, string bodyId, TasksDelegate task)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var tasksDAL = new TasksDelegateDALprovider(eventDAL);

                var taskDTO = await MapUpdate(bodyCtx, task);
                await tasksDAL.AddorUpdateForEvent(bodyCtx, taskDTO.IdMeeting, taskDTO);

                // 2. Actions to be triggered 
                switch (task.TaskStatusId)
                {
                    case TaxonomyValuesIds.RequestStatus.Accepted:
                        //Tras la aceptación se se debe generar una Tarea de Aceptación para el Delegado
                        await NewDelegationApprodval(bodyCtx, bodyId, taskDTO);
                        break;

                    case TaxonomyValuesIds.RequestStatus.Rejected:
                        //Actualizar la petición Padre a Pendiente
                        await PendingTaskAttendance(bodyCtx, bodyId, taskDTO);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateTaskDelegate by {bodyId} with taskAsistance: {task}", ex);
            }
        }

        private async Task UpdateAttendanceTask(IPnPContext bodyCtx, string bodyId, TasksDelegate taskDelegate)
        {
            var tasksService = _serviceprovider.GetRequiredService<TasksAttendanceService>();

            //Se actualiza tarea Padre a Pendiente Delegation
            var upAttendance = new TasksAttendance()
            {
                TaskId = taskDelegate.TaskParentId,
                TaskStatusId = DALConstants.TaxonomyValuesIds.RequestStatus.PendingDelegation,
                Vote = null,
                HasAttended = null,
            };

            await tasksService.UpdateTask(bodyCtx, bodyId, upAttendance);
        }

        private async Task NewDelegationApprodval(IPnPContext ctx, string bodyId, TasksDelegateDTO tDTO)
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = _serviceprovider.GetRequiredService<TasksAttendanceService>();

            // 1. La Tarea se queda en estado Delegada
            var updateTask = new TasksAttendance()
            {
                TaskId = tDTO.IdRequestPadre,
                TaskStatusId = TaxonomyValuesIds.RequestStatus.Delegated,
                DelegateTo = tDTO.Delegado.AsUserPrincipalName() ?? string.ECNTy,
                DelegatedUserVote = tDTO.DelegadoVoto.AsUserPrincipalName() ?? string.ECNTy,
                DelegateVote = tDTO.VotoDelegado,
            };
            await tasksDAL.UpdateTask(ctx, bodyId, updateTask);

            // 2.1. Se agrega al delegado de Attendance a la Meeting (si todavia no esta añadido)
            var storedEvent = await eventDAL.GetBySharedEventId(ctx, tDTO.IdMeeting);
            var exist = storedEvent?.Asistentes?.Any(t => t.LookupId == tDTO?.Delegado?.LookupId) ?? false;
            if (!exist)
            {
                var asistentes = storedEvent.Asistentes.ToList();
                asistentes.Add(tDTO.Delegado);
                storedEvent.Asistentes = asistentes.ToArray();

                await storedEvent.AsListItem().UpdateAsync();
            }

            // 2.2. Se agrega al delegado de voto a la Meeting (si todavia no esta añadido)
            exist = storedEvent?.Asistentes?.Any(t => t.LookupId == tDTO?.DelegadoVoto?.LookupId) ?? false;
            if (!exist)
            {
                var asistentes = storedEvent.Asistentes.ToList();
                asistentes.Add(tDTO.Delegado);
                storedEvent.Asistentes = asistentes.ToArray();

                await storedEvent.AsListItem().UpdateAsync();
            }

            // 3. Se notifica la Delegation de Attendance y/o voto
            var url = _bodyRoleService.GetNotificationUrlByEvent(ctx, tDTO.IdMeeting, bodyId);
            var taskUser = await tDTO.SolicitadoPor.AsSPUser(ctx);
            var valuesReplace = new Dictionary<string, string>()
            {
                {ReplacedTokens.MeetingTitle, storedEvent.Title},
                {ReplacedTokens.DelegatedBy, taskUser?.Title ?? ""}
            };

            if (tDTO.Delegado.LookupId == tDTO.DelegadoVoto.LookupId)
            {
                // Se notifica la Delegation de Attendance y voto al mismo usuario
                await _notificationsService.AddUserNotification(bodyId, storedEvent.MeetingType!.TermId, tDTO.Delegado.AsUserPrincipalName(), NotificationsMessages.Alert.InformacionDelegationAttendanceYVoto, _defaultLocale, url, valuesReplace);
            }
            else
            {
                // Se notifica la Delegation de Attendance a un usuario
                await _notificationsService.AddUserNotification(bodyId, storedEvent.MeetingType!.TermId, tDTO.Delegado.AsUserPrincipalName(), NotificationsMessages.Alert.InformacionDelegationAttendance, _defaultLocale, url, valuesReplace);

                // Se notifica la Delegation de voto a otro usuario
                await _notificationsService.AddUserNotification(bodyId, storedEvent.MeetingType!.TermId, tDTO.DelegadoVoto.AsUserPrincipalName(), NotificationsMessages.Alert.InformacionDelegationVoto, _defaultLocale, url, valuesReplace);
            }
        }

        private async Task PendingTaskAttendance(IPnPContext ctx, string bodyId, TasksDelegateDTO task)
        {


            var taskAsistance = new TasksAttendance()
            {
                TaskId = task.IdRequestPadre,
                TaskStatusId = TaxonomyValuesIds.RequestStatus.Pending,
            };

            await _tasksAttendanceService.UpdateTask(ctx, bodyId, taskAsistance);
            await NotifyRejectDelegation(ctx, bodyId, task);
        }

        private async Task NotifyRejectDelegation(IPnPContext ctx, string bodyId, TasksDelegateDTO task)
        {
            try
            {
                var eventDAL = new InConstructionEventsDALprovider();
                var attendanceprovider = new TasksAttendanceDALprovider(eventDAL);

                var ev = await eventDAL.GetBySharedEventId(ctx, task.IdMeeting);
                var parentTask = await attendanceprovider.GetTaskById(ctx, int.Parse(task.IdRequestPadre));

                var url = _bodyRoleService.GetNotificationUrlByEvent(ctx, task.IdMeeting, bodyId);
                var valuesReplace = new Dictionary<string, string>()
            {
                {ReplacedTokens.MeetingTitle, ev.Title},
                {ReplacedTokens.RejectionReason, task.Comentario}
            };

                await _notificationsService.AddUserNotification(bodyId, ev.MeetingType!.TermId, parentTask.AsignadoA.AsUserPrincipalName(), NotificationsMessages.Alert.RechazarDelegation, _defaultLocale, url, valuesReplace);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error to send reject notification with task id {task.Id} from {bodyId}", ex);
            }
        }

        private async Task<TasksDelegateDTO> MapAddDelegate(IPnPContext ctx, TasksDelegate item)
        {
            TasksDelegateDTO del = null;
            if (item != null)
            {
                del = new TasksDelegateDTO
                {
                    Title = item.TaskTitle,
                    EstadoRequest = item.TaskStatusId,
                    IdRequestPadre = item.TaskParentId,
                    Comentario = item.Comment,
                };

                if (item.DelegateVote != null)
                    del.VotoDelegado = item.DelegateVote.GetValueOrDefault();

                if (!String.IsNullOrECNTy(item.TaskId))
                    del.Id = int.Parse(item.TaskId);

                if (item.TaskStartDate != null)
                    del.RequestStartDate = item.TaskStartDate;

                if (item.TaskEndDate != null)
                    del.RequestEndDate = item.TaskEndDate;

                // ensure all users are testsent in the site
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
                del.DelegadoVoto = !string.IsNullOrECNTy(item.DelegatedUserVote)
                   && ensuredUsers.ContainsKey(item.DelegatedUserVote) ?
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

        private async Task<TasksDelegateDTO> MapUpdate(IPnPContext ctx, TasksDelegate item)
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var delegateDAL = new TasksDelegateDALprovider(eventsprovider);

            // Update Task     

            TasksDelegateDTO? taskSave = null;
            await RunAsSystem(ctx, async (ctxSystem) =>
             {
                 taskSave = await delegateDAL.GetTaskById(ctxSystem, int.Parse(item.TaskId));
             });

            if (taskSave is not null)
            {
                if (!String.IsNullOrECNTy(item.TaskParentId))
                    taskSave.IdRequestPadre = item.TaskParentId;

                if (!String.IsNullOrECNTy(item.TaskStatusId))
                    taskSave.EstadoRequest = item.TaskStatusId;

                if (!String.IsNullOrECNTy(item.TaskTitle))
                    taskSave.Title = item.TaskTitle;

                if (!String.IsNullOrECNTy(item.Comment))
                    taskSave.Comentario = item.Comment;

                // ensure all users are testsent in the site
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

                if (!String.IsNullOrECNTy(item.RequestedBy))
                    taskSave.SolicitadoPor = new FieldUserValue(ensuredUsers[item.RequestedBy]);
                if (!String.IsNullOrECNTy(item.DelegateTo))
                    taskSave.Delegado = new FieldUserValue(ensuredUsers[item.DelegateTo]);
                if (!String.IsNullOrECNTy(item.DelegatedUserVote))
                    taskSave.Delegado = new FieldUserValue(ensuredUsers[item.DelegatedUserVote]);

                if (item.AssignedTo != null && item.AssignedTo.Any())
                    taskSave.AsignadoA = (item.AssignedTo ?? new List<string>())
                        .Where(u => ensuredUsers.ContainsKey(u))
                        .Select(u => new FieldUserValue(ensuredUsers[u]))
                        .ToArray();

                if (item.TaskStartDate != null)
                    taskSave.RequestStartDate = item.TaskStartDate;

                if (item.TaskEndDate != null)
                    taskSave.RequestEndDate = item.TaskEndDate;

                if (item.DelegateVote != null)
                    taskSave.VotoDelegado = item.DelegateVote.GetValueOrDefault();
            }


            return taskSave;
        }

        private void CheckAddTaskDelegate(TasksDelegate t)
        {
            if (String.IsNullOrECNTy(t.SharedEventId))
                throw new ArgumentNullException(nameof(t.SharedEventId));

            if (String.IsNullOrECNTy(t.TaskParentId))
                throw new ArgumentNullException(nameof(t.TaskParentId));

            if (String.IsNullOrECNTy(t.RequestedBy))
                throw new ArgumentNullException(nameof(t.RequestedBy));

            if (String.IsNullOrECNTy(t.DelegateTo))
                throw new ArgumentNullException(nameof(t.DelegateTo));
        }

        public async Task CancelOrExpiredCertification(IPnPContext bodyCtx, string bodyId, IEnumerable<TasksDTO> taskDelegation, string termId)
        {
            try
            {
                foreach (var delegation in taskDelegation)
                {
                    var task = new TasksDelegate()
                    {
                        TaskId = delegation.Id.ToString(),
                        TaskStatusId = termId
                    };
                    var isEditorOrSystem = await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(bodyCtx, bodyId);
                    var taskDTO = await MapUpdate(bodyCtx, task);
                    await SaveTaskDelegation(bodyCtx, bodyId, taskDTO);
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"Error Cancel or expired delegation task by {bodyId}", ex);
            }
        }

        private async Task SaveTaskDelegation(IPnPContext bodyCtx, string bodyId, TasksDelegateDTO taskDTO)
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var tasksDAL = new TasksDelegateDALprovider(eventDAL);
            await tasksDAL.AddorUpdateForEvent(bodyCtx, taskDTO.IdMeeting, taskDTO);
        }
    }
}