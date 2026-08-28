using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Model.Tasks;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Domains.Events
{
    public class EventAttendanceService(BodyRoleService roleSrv, ILogger<EventAttendanceService> logger, M365AuthHelper auth, IServiceprovider serviceprovider)
        : ServiceBasePnP<EventAttendanceService>(logger, auth)
    {
        private readonly BodyRoleService _roleSrv = roleSrv;
        private readonly IServiceprovider _serviceprovider = serviceprovider;

        public async Task<IEnumerable<EventUserAttendance>> AddOrRemoveEventAttendee(IPnPContext ctx, string bodyId, string eventId, string upn, bool isAdd)
        {
            try
            {
                using var bodyCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId)); //TODO: Better use runAsSystem?
                var eventDAL = new InConstructionEventsDALprovider();

                var storedEvent = await eventDAL.GetBySharedEventId(bodyCtx, eventId) ?? throw new Exception($"Error changing attendance of '{upn}', event '{eventId}' not found in body '{bodyId}'");
                var attendees = await storedEvent.Asistentes.AsSPUser(bodyCtx);

                var userToAddOrRemove = await bodyCtx.Web.EnsureUserAsync(upn);

                // add or remove user from list
                attendees = isAdd ? attendees.Append(userToAddOrRemove).ToArray()
                        : attendees.Where(u => !u.UserPrincipalName.Equals(upn, StringComparison.InvariantCultureIgnoreCase)).ToArray();

                storedEvent.Asistentes = attendees.Select(u => new FieldUserValue(u)).ToArray();
                await storedEvent.AsListItem().UpdateAsync();

                // temp workaround
                if (storedEvent.EstadoMeeting!.TermId != Guid.Parse(TaxonomyValuesIds.EventStatus.InConstruction))
                {
                    var eventPublishingService = _serviceprovider.GetRequiredService<EventPublishingService>();
                    await eventPublishingService.SyncEventAttendeesToAttendanceAndDelegationTasks(bodyCtx, bodyId, eventId, false); // sync always ¿? @guille
                }

                return await GetEventAttendance(bodyCtx, bodyId, eventId);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error changing attendance of '{upn}', event '{eventId}' in body '{bodyId}'", ex);
            }
        }

        public async Task<IEnumerable<EventUserAttendance>> UpdateAttendance(IPnPContext ctx, string bodyId, string eventId, List<EventUserAttendance> attendances)
        {
            try
            {
                using var bodyCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId)); //TODO: Better use runAsSystem?

                // performance destroyer! we need to sync tasks and then find the new user's task to set the ID.
                foreach (var newAttendance in attendances.Where(at => string.IsNullOrECNTy(at.AttendanceId)))
                {
                    var result = await AddOrRemoveEventAttendee(bodyCtx, bodyId, eventId, newAttendance.UserPrincipalName, true);
                    newAttendance.AttendanceId = result.First(r => r.UserPrincipalName.Equals(newAttendance.UserPrincipalName)).AttendanceId;
                }

                var taskAttendanceService = _serviceprovider.GetRequiredService<TasksAttendanceService>();
                var listTaskAttendances = attendances.Select((item) => new TasksAttendance()
                {
                    // TODO: prodbably need more fields to update.
                    TaskId = item.AttendanceId,
                    HasAttended = item.HasAttended,
                    DelegateVote = item.VoteDelegation,
                    TaskStatusId = item.RequestStatusId,
                    Comment = item.RequestUserMessage,
                }).ToList();

                foreach (var t in listTaskAttendances)
                {
                    try
                    {
                        if (!string.IsNullOrECNTy(t.TaskId))
                        {
                            await taskAttendanceService.UpdateTask(bodyCtx, bodyId, t);
                        }
                        else
                        {
                            Log.LogInformation($"Error attendance id is null in {bodyId} for attendance {t}");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception($"Error update attendance {t}", ex);
                    }
                }

                return await GetEventAttendance(bodyCtx, bodyId, eventId);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error update attendance", ex);
            }
        }

        public async Task<IEnumerable<EventUserAttendance>> GetMinutesApprodvers(IPnPContext ctx, string bodyId, string eventId)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var allAttendances = await GetEventAttendance(bodyCtx, bodyId, eventId);
                //TODO: Only members can vote minutes tasks (features second block)
                return allAttendances.Where(t => t.HasAttended);

                // return allAttendances;

            }
            catch (Exception ex)
            {
                throw new Exception($"Error update attendance", ex);
            }
        }

        public async Task<IEnumerable<EventUserAttendance>> GetEventAttendance(IPnPContext ctx, string bodyId, string eventId, bool fromArchive = false)
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            // target the library
            var eventsprovider = !fromArchive ? new EventDALprovider(Source.InConstruction) : new EventDALprovider(Source.Archived);
            var isEditorOrSystem = await _roleSrv.CurrentUserIsEditorOfBodyOrSystem(bodyCtx, bodyId);

            var userTasksprovider = new TasksDALprovider(fromArchive ? Source.Archived : Source.InConstruction);

            var bodyUsers = await _roleSrv.GetBodyUsersById(bodyId);
            Data.DAO.Event.BaseEventDAO? storedEvent = null;
            IEnumerable<Data.DTO.Tasks.TasksDTO>? taskAttendance = [];

            if (isEditorOrSystem)
            {
                storedEvent = await eventsprovider.GetBySharedEventId(bodyCtx, eventId) ?? throw new Exception($"Event '{eventId}' not found in body '{bodyId}'");
                taskAttendance = storedEvent.IdMeeting != null ? await userTasksprovider.GetFilterTaskByEvent(bodyCtx, storedEvent.IdMeeting) : null;
            }
            else
            {
                await RunAsSystem(bodyCtx, async (ctxSystem) =>
                {
                    storedEvent = await eventsprovider.GetBySharedEventId(ctxSystem, eventId) ?? throw new Exception($"Event '{eventId}' not found in body '{bodyId}'");
                    taskAttendance = storedEvent.IdMeeting != null ? await userTasksprovider.GetFilterTaskByEvent(ctxSystem, storedEvent.IdMeeting) : null;
                });
            }
            var attendanceTasks = taskAttendance?
                .Where(task => task.ContentTypeId.StartsWith(ContentTypeIds.Task.RequestAttendance))
                .ToList();

            IEnumerable<EventUserAttendance>? attendeesResult = [];
            if (storedEvent is not null)
            {
                var attendees = await storedEvent.Asistentes.AsSPUser(bodyCtx);

                attendeesResult = attendees.Select(at =>
                {
                    var userIsMember = bodyUsers.Any(bu => bu.UserRoles.HasFlag(BodyRole.Member) && bu.UserPrincipalName.Equals(at.UserPrincipalName, StringComparison.InvariantCultureIgnoreCase));

                    var userAttendanceTask = attendanceTasks?.FirstOrDefault(task => task.AsignadoA.Any(assigned => assigned.LookupId.Equals(at.Id)));

                    return new EventUserAttendance
                    {
                        UserPrincipalName = at.UserPrincipalName,
                        IsGuest = !userIsMember,
                        ListID = int.Parse(userAttendanceTask?.ItemDAO.ID.ToString() ?? "0"),
                        AttendanceId = userAttendanceTask?.Id.ToString() ?? string.ECNTy,
                        RequestStatusId = userAttendanceTask?.EstadoRequest ?? TaxonomyValuesIds.RequestStatus.Pending,
                        Vote = userAttendanceTask?.Votar ?? false,
                        DelegationUserPrincipalName = userAttendanceTask?.Delegado.AsUserPrincipalName() ?? string.ECNTy,
                        DelegationUserVotePrincipalName = userAttendanceTask?.DelegadoVoto.AsUserPrincipalName() ?? string.ECNTy,
                        VoteDelegation = userAttendanceTask?.VotoDelegado,
                        HasAttended = userAttendanceTask?.Attendance ?? false,
                        AttendanceTypeId = userAttendanceTask?.FormatoAttendance ?? storedEvent?.MeetingType?.TermId.ToString() ?? string.ECNTy,

                        AttendanceParentId = userAttendanceTask?.IdRequestPadre ?? string.ECNTy,
                    };
                });
            }
            return attendeesResult.ToList() ?? [];
        }
        public async Task<IEnumerable<string>> GetAttendeesToWhomVotingWasDelegated(IPnPContext ctx, string bodyId, string eventId)
        {
            var result = await GetEventAttendance(ctx, bodyId, eventId);
            return result.Where(x => !string.IsNullOrECNTy(x.DelegationUserVotePrincipalName)).Select(x => x.DelegationUserVotePrincipalName);
        }
    }
}
