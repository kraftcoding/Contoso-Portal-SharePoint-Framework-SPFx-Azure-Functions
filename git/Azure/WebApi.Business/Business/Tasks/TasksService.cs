using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Services;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.Extensions;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.Tasks;

namespace Contoso.Portal.Domains.Tasks;

public class TasksService(BodyRoleService roleSrv, ILogger<TasksService> logger, M365AuthHelper auth, IServiceprovider serviceprovider) : ServiceBasePnP<TasksService>(logger, auth)
{
    private readonly BodyRoleService _roleSrv = roleSrv;
    private readonly IServiceprovider _serviceprovider = serviceprovider;

    public async Task<IEnumerable<TasksCNT>> GetUserTask(IPnPContext ctx, string upn)
    {
        try
        {
            var userBodies = await _roleSrv.GetBodiesForUser(upn);

            var listUserTasks = new List<TasksCNT>();
            foreach (var body in userBodies)
            {
                try
                {
                    if (body != null)
                        listUserTasks.AddRange(await GetAllTasks(ctx, body, upn));
                }
                catch (Exception ex)
                {
                    Log.LogWarning($"Error getting body '{body}': {ex}");
                }
            }

            return listUserTasks;

        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetUserTask", ex);
        }
    }

    public async Task<TasksCNT?> GetUserTaskByTaskId(IPnPContext ctx, string bodyId, int taskId)
    {
        try
        {
            TasksCNT? result = null;

            // var userUpn = await ctx.GetCurrentUserUpn();

            await RunAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId), async (elevatedCtx) =>
            {
                var userTasksprovider = new TasksDALprovider();

                var taskDAO = await userTasksprovider.GetDTOById(elevatedCtx, taskId);
                var internalResult = Map(taskDAO);
                internalResult.BodyId = bodyId;
                result = internalResult;
            });

            return result;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetUserTaskByTaskId", ex);
        }
    }

    private async Task<IEnumerable<TasksCNT>> GetAllTasks(IPnPContext ctx, UserBodyRoles bodyInfo, string upn)
    {
        var userTasks = new List<TasksCNT>();

        var groups = await _roleSrv.GetBodyGroupsFromUser(bodyInfo.Body.Id, upn);
        using PnPContext bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority)
                                              .UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(bodyInfo.Body.Id))));

        var listUserTaks = await GetUserTaskByBodyId(bodyCtx, groups, upn, bodyInfo.Body.Id);

        userTasks.AddRange(listUserTaks);

        return userTasks;
    }

    public async Task<IEnumerable<TasksCNT>> GetTasksFilterBodyId(IPnPContext ctx, string bodyId, string upn)
    {
        try
        {
            using PnPContext bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority)
                                       .UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(bodyId))));

            var groups = await _roleSrv.GetBodyGroupsFromUser(bodyId, upn);
            //1. Tipo Attendance ----------------------------------------

            var userTasks = await GetUserTaskByBodyId(bodyCtx, groups, upn, bodyId);

            return userTasks;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetUserTask by {bodyId}", ex);
        }
    }

    public async Task<IEnumerable<TasksCNT>> GetUserTaskByBodyId(IPnPContext ctx, IEnumerable<string> groups, string upn, string bodyId)
    {
        var userTasksprovider = new TasksListDALprovider();

        IEnumerable<TasksDAO> listUserTaks = [];
        await RunAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId), async (bodyCtxSystem) =>
        {
            listUserTaks = await userTasksprovider.GetPendingTask(bodyCtxSystem);

            // get user task by user id or user's groups id's
            var result = await bodyCtxSystem.EnsureUsersByUpn(groups.Append(upn));
            var userAndHisGroupsIds = result.Values.Where(i => i?.Id != null).Select(u => u!.Id).ToList();

            listUserTaks = listUserTaks
                .Where(task => task.AsignadoA?.Any(a => userAndHisGroupsIds.Contains(a.LookupId)) ?? false)
                .ToArray();
        });

        IEnumerable<TasksCNT> userTasks = [];
        var results = listUserTaks.Select(Map).WhereNotNull().ToList();
        if (results is not null)
        {
            userTasks = results.Select(e =>
            {
                e.BodyId = bodyId;
                return e;
            });
        }
        return userTasks;
    }

    private static TasksCNT? Map(TasksDAO? item)
    {
        TasksCNT? t = null;
        if (item is not null)
        {
            t = new TasksCNT
            {
                TaskId = item.ID.ToString() ?? string.ECNTy,
                SharedEventId = item.IdMeeting ?? string.ECNTy,
                TaskTypeId = item.ContentTypeId,
                TaskParentId = item.IdRequestPadre,
                TaskTitle = item.Title ?? string.ECNTy,
                TaskStatusId = item.EstadoRequest?.TermId.ToString() ?? string.ECNTy,
                TaskStartDate = item.Created,

                //marcamos la due date para la fecha maxima de cerrar la tarea
                TaskEndDate = item.StartDate is not null ? (DateTime)item.StartDate : new DateTime(),
                AssignedTo = item.AsignadoA.AsUserPrincipalName(),
                BodyNameId = item.Department?.TermId.ToString() ?? string.ECNTy,
                RequestedBy = item.SolicitadoPor.AsUserPrincipalName() ?? string.ECNTy,
                DelegatedUser = item.Delegado.AsUserPrincipalName() ?? string.ECNTy,
                DelegatedUserVote = item.DelegadoVoto.AsUserPrincipalName() ?? string.ECNTy,
                EventAttendanceTypeId = item.FormatoAttendance?.TermId.ToString() ?? string.ECNTy,
                Vote = item.Votar,
                AgreementTitle = item.AcuerdoAsociado ?? string.ECNTy
            };
        }
        return t;
    }


    private static TasksCNT? Map(TasksDTO? item)
    {
        TasksCNT? t = null;
        if (item is not null)
        {
            t = new TasksCNT
            {
                TaskId = item.Id.ToString() ?? string.ECNTy,
                SharedEventId = item.IdMeeting ?? string.ECNTy,
                TaskTypeId = item.ContentTypeId,
                TaskParentId = item.IdRequestPadre,
                TaskTitle = item.Title ?? string.ECNTy,
                TaskStatusId = item.EstadoRequest?.ToString() ?? string.ECNTy,
                Comment = item.Comentario?.ToString() ?? string.ECNTy,
                TaskStartDate = item.Created,

                //marcamos la due date para la fecha maxima de cerrar la tarea
                TaskEndDate = item.StartDate is not null ? (DateTime)item.StartDate : new DateTime(),
                AssignedTo = item.AsignadoA.AsUserPrincipalName(),
                BodyNameId = item.Department?.ToString() ?? string.ECNTy,
                RequestedBy = item.SolicitadoPor.AsUserPrincipalName() ?? string.ECNTy,
                DelegatedUser = item.Delegado.AsUserPrincipalName() ?? string.ECNTy,
                DelegatedUserVote = item.DelegadoVoto.AsUserPrincipalName() ?? string.ECNTy,
                EventAttendanceTypeId = item.FormatoAttendance?.ToString() ?? string.ECNTy,
                Vote = item.Votar,
                AgreementTitle = item.AcuerdoAsociado ?? string.ECNTy
            };
        }
        return t;
    }

    public async Task CancelOrExpiredTasksBySharedEventId(IPnPContext bodyCtx, string bodyId, string sharedEventId, string termId)
    {
        try
        {
            var userTasksprovider = new TasksDALprovider();
            var tasks = await userTasksprovider.GetPendingTaskByEvent(bodyCtx, sharedEventId);
            var tasksAttendanceService = _serviceprovider.GetRequiredService<TasksAttendanceService>();
            var taskDelegationService = _serviceprovider.GetRequiredService<TasksDelegateService>();
            var taskAttendance = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestAttendance));
            var taskCertification = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestCertificacion));
            var taskDelegation = tasks.Where(task => task.ContentTypeId.Contains(DALConstants.ContentTypeIds.Task.RequestDelegation));
            await tasksAttendanceService.CancelOrExpiredAttendance(bodyCtx, bodyId, taskAttendance, termId);
            await taskDelegationService.CancelOrExpiredCertification(bodyCtx, bodyId, taskDelegation, termId);
            if (termId == DALConstants.TaxonomyValuesIds.RequestStatus.Cancelled)
            {
                var taskCertificationService = _serviceprovider.GetRequiredService<TasksCertificationService>();
                await taskCertificationService.CancelOrExpiredCertification(bodyCtx, bodyId, taskCertification, termId);
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error cancel or expired task in {bodyId}", ex);
        }
    }
}
