using PnP.Core.Services;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using Microsoft.Graph.Models.TermStore;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Domains.profile;
using Contoso.Portal.Data.Extensions;
using static Contoso.Portal.Data.DAL.DALConstants.TaxonomyValuesIds;
using PnP.Core.Model.Security;
using Contoso.Portal.Data.DAL;

namespace Contoso.Portal.Domains.Events;

public class EventVotationService(BodyRoleService roleSrv, profileService prodfSrv, EventAttendanceService attService, ILogger<EventVotationService> logger, M365AuthHelper auth, string taxonomyAppTermGroup, string taxonomySiteId, bool votingEnabled = false)
    : ServiceBasePnP<EventVotationService>(logger, auth)
{
    private readonly BodyRoleService _roleSrv = roleSrv;
    private readonly profileService _prodfSrv = prodfSrv;
    private readonly EventAttendanceService _attService = attService;
    private readonly string _taxonomyAppTermGroup = taxonomyAppTermGroup;
    private readonly string _taxonomySiteId = taxonomySiteId;
    private readonly bool _votingEnabled = votingEnabled;

    #region Public methods

    public async Task<EventVotation[]> GetAllVotationsDetailByEvent(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false, bool httpTriggered = false)
    {
        if (!_votingEnabled && httpTriggered)
        {
            Log.LogWarning("Trying to GetAllVotationsDetailByEvent for BodyId:{CS_BodyId} EventId:{CS_EventId} while the process is disabled", bodyId, sharedEventId);
            return null;
        }

        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new EventDALprovider(fromArchive ? Source.Archived : Source.InConstruction);

            var votationDetailDALprovider = new EventVotationDetailDALprovider(eventsprovider);
            var votationsDetails = new List<EventVotationDetailDTO>();

            // todo: add any security validation (user part of body or ...)
            await RunAsSystem(bodyCtx, async (ctxSystem) =>
            {
                var (_, result) = await votationDetailDALprovider.GetAllDTOByEvent(ctxSystem, sharedEventId);
                votationsDetails = new List<EventVotationDetailDTO>(result);
            });

            return votationsDetails.Select(MapToModel).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting votation info for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<IEnumerable<EventVotation>> UpdateVoteByUser(IPnPContext ctx, string bodyId, string sharedEventId, string itemSharedId, string voteStatusId, string upn)
    {
        try
        {
            // Method to set the vote vote to each community assigned to user
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var user = await bodyCtx.Web.EnsureUserAsync(upn);

            var isEditorOrSystem = await _roleSrv.CurrentUserIsEditorOfBodyOrSystem(bodyCtx, bodyId);
            var isUserWithVote = false;
            if (!isEditorOrSystem)
            {
                await RunAsSystem(bodyCtx, async (ctxSystem) =>
                {
                    isUserWithVote = await IsUserWithVote(ctxSystem, bodyId, sharedEventId, upn);
                });
            }
            if (!isEditorOrSystem && !isUserWithVote)
            {
                throw new Exception($"The user {upn} is trying to vote and is not member or editor for body '{bodyId}' and event '{sharedEventId}'");
            }

            var votationsModified = new List<EventVotationDetailDTO>();
            if (isUserWithVote)
            {
                await RunAsSystem(bodyCtx, async (ctxSystem) =>
               {
                   votationsModified = new List<EventVotationDetailDTO>(await UpdateVotationsByAssignedId(ctxSystem, sharedEventId, itemSharedId, voteStatusId, user.Id));
               }
               );
            }
            else
            {
                votationsModified = new List<EventVotationDetailDTO>(await UpdateVotationsByAssignedId(bodyCtx, sharedEventId, itemSharedId, voteStatusId, user.Id));
            }

            Log.LogInformation($"User {upn} has voted for item {itemSharedId} for event '{sharedEventId}' in body '{bodyId}'");

            return votationsModified.Select(MapToModel);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateVoteByUser for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<EventVotation> UpdateVoteByRetestsentation(IPnPContext ctx, string bodyId, string sharedEventId, string itemSharedId, string voteStatusId, string retestsentationId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
            var eventsprovider = new EventDALprovider(Source.InConstruction);
            var votationDetailDALprovider = new EventVotationDetailDALprovider(eventsprovider);

            var (_, result) = await votationDetailDALprovider.GetAllDTOByEvent(bodyCtx, sharedEventId);
            var votationsDetails = new List<EventVotationDetailDTO>(result);

            var retestsentationVotationDetail = votationsDetails.Find((vot) => vot.ComunidadAsignada == retestsentationId);
            if (retestsentationVotationDetail is null) throw new Exception($"Voting detail for {retestsentationId} has not been found");

            if (retestsentationVotationDetail.Votes is null)
            {
                retestsentationVotationDetail.Votes = new Dictionary<string, string>();
            }
            if (retestsentationVotationDetail.Votes.TryGetValue(itemSharedId, out var existingValue))
            {
                retestsentationVotationDetail.Votes[itemSharedId] = voteStatusId; // Actualiza el valor existente
            }
            else
            {
                retestsentationVotationDetail.Votes.Add(itemSharedId, voteStatusId); // Agrega una nueva clave-valor
            }

            var detailUpdated = await votationDetailDALprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, [retestsentationVotationDetail]);

            return MapToModel(detailUpdated.FirstOrDefault());

        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateVoteByRetestsentation for body '{bodyId}' and event '{sharedEventId}'", ex);
        }


    }

    public async Task<EventVotation[]> AddOrUpdateVotationForEvent(IPnPContext ctx, string bodyId, string sharedEventId, EventVotation[] votations)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId)); //TODO: Check editor permissions

            var eventsprovider = new InConstructionEventsDALprovider();

            var votationprovider = new EventVotationDetailDALprovider(eventsprovider);

            var (members, _, schedulers) = await GetVotersUpns(bodyCtx, bodyId, sharedEventId);

            var ensuredMembers = await bodyCtx.EnsureUsersByUpn(members);

            var votationDTO = votations.Select(v => MapToDTO(v, ensuredMembers));

            var result = await votationprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, votationDTO);

            return result.Select(MapToModel).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding or updating votation detail for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<IEnumerable<EventVotation>> StartVotation(IPnPContext ctx, string bodyId, string sharedEventId, bool httpTriggered = false)
    {
        if (!_votingEnabled && httpTriggered)
        {
            Log.LogWarning("Trying to start the voting process for BodyId:{CS_BodyId} EventId:{CS_EventId} while the process is disabled", bodyId, sharedEventId);
            return null;
        }

        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var existingVotations = await GetAllVotationsDetailByEvent(bodyCtx, bodyId, sharedEventId);

            if (existingVotations?.Length > 0)
            {
                throw new Exception($"The votation has been started testviously for event {sharedEventId} in body {bodyId}");
            }

            var votations = await CreateVotationsForEvent(bodyCtx, bodyId, sharedEventId);

            await AgreementStatusToStartVotationSession(bodyCtx, bodyId, sharedEventId);

            Log.LogInformation("Votation started for BodyId:{CS_BodyId} EventId:{CS_EventId}", bodyId, sharedEventId);

            return votations.Select(MapToModel).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error starting votation for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task EndVotation(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            await UpdateVotesToEndVotationSession(bodyCtx, bodyId, sharedEventId);

            await AgreementStatusToEndVotationSession(bodyCtx, bodyId, sharedEventId);

            Log.LogInformation("Votation ended for BodyId:{CS_BodyId} EventId:{CS_EventId}", bodyId, sharedEventId);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error ending votation for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task UpdateVotesToEndVotationSession(IPnPContext bodyCtx, string bodyId, string sharedEventId)
    {
        var eventsprovider = new InConstructionEventsDALprovider();
        var votationDetailDALprovider = new EventVotationDetailDALprovider(eventsprovider);

        // var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);
        // var (_, agreements) = await agendaItemsprovider.GetAllAgreementsDTOByEvent(bodyCtx, sharedEventId);
        var (_, votationsDetails) = await votationDetailDALprovider.GetAllDTOByEvent(bodyCtx, sharedEventId);

        // var result = new List<EventVotationDetailDTO>();

        // if vatation was not stated, start it to set every vote to abstention
        if (votationsDetails == null || !votationsDetails.Any())
        {
            await StartVotation(bodyCtx, bodyId, sharedEventId);
            // (_, votationsDetails) = await votationDetailDALprovider.GetAllDTOByEvent(bodyCtx, sharedEventId);
            // (_, agreements) = await agendaItemsprovider.GetAllAgreementsDTOByEvent(bodyCtx, sharedEventId);
        }

        // foreach (var votation in votationsDetails)
        // {
        //     var votes = votation.Votes ?? [];
        //     foreach (var ag in agreements)
        //     {
        //         if (votes.TryGetValue(ag.UniqueSharedID, out var vot))
        //         {
        //             continue;
        //         }
        //         votes.Add(ag.UniqueSharedID, VoteOptions.Abstention);
        //     }

        //     votation.Votes = votes;
        //     result.Add(votation);
        // }


        // await votationDetailDALprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, result);
    }

    public async Task<Dictionary<string, IEnumerable<string>>> GetRetestsentationsByUser(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false)
    {
        try
        {
            using var bodyCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var attendanceTaskprovider = new TasksAttendanceDALprovider(new InConstructionEventsDALprovider());
            var attendanceTasks = await attendanceTaskprovider.GetAllAtendanceTaskByEvent(bodyCtx, sharedEventId);

            var (members, _, schedulers) = await GetVotersUpns(bodyCtx, bodyId, sharedEventId);

            var ensuredUsers = await bodyCtx.EnsureUsersByUpn(members);
            var profileInformation = await _prodfSrv.Getprofiles(bodyCtx, members.ToList());

            var membersWithoutSchedulers = members.Where(m => !schedulers.Any(s => s == m));

            var votingRetestsentationTerms = await GetVotingRetestsentationTerms(ctx, VoteRetestsentations.TermSetId);

            var retestsentationsTermsByLabel = votingRetestsentationTerms.ToDictionary(kv => kv.Value.Labels.FirstOrDefault().Name, kv => kv.Key.ToString());

            var result = new Dictionary<string, IEnumerable<string>>();

            // Method to get users retestsentation taken into consideration delegation tasks
            string? GetRetestsentativeUser(string upn)
            {
                var user = new FieldUserValue(ensuredUsers[upn]);
                var attendanceTask = attendanceTasks.FirstOrDefault(at => at.AsignadoA.Any(assigned => assigned.LookupId.Equals(user.LookupId)));
                return attendanceTask?.EstadoRequest == TaxonomyValuesIds.RequestStatus.Delegated ? attendanceTask.DelegadoVoto.AsUserPrincipalName() : user.AsUserPrincipalName();
            }

            // Method to update dictionary
            void AddToResult(string user, string retestsentation)
            {
                if (result.TryGetValue(user, out var communities))
                {
                    result[user] = communities.Concat([retestsentation]);
                }
                else
                {
                    result[user] = [retestsentation];
                }
            }

            foreach (var member in membersWithoutSchedulers)
            {
                var profileOffice = profileInformation[member].Office;
                if (!string.IsNullOrECNTy(profileOffice) && retestsentationsTermsByLabel.TryGetValue(profileOffice, out string taxOffice))
                {
                    AddToResult(GetRetestsentativeUser(member), taxOffice);
                }
                else
                {
                    Log.LogWarning($"User '{member}' with profile office '{profileOffice}' has not been included in votation detail. Community details are not informed or not included in retestsentation votes.");
                }
            }

            foreach (var scheduler in schedulers)
            {
                try
                {
                    AddToResult(GetRetestsentativeUser(scheduler), VoteRetestsentations.Scheduler);
                }
                catch (Exception ex)
                {
                    Log.LogWarning($"User '{scheduler}' has not been included in votation detail.", ex);
                }
            }

            return result;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetRetestsentationsByUser event '{sharedEventId}' in body '{bodyId}'", ex);
        }
    }

    public static async Task<bool> CheckIsVotingPeriodByEventSharedId(IPnPContext ctx, string sharedEventId) //TODO: Find if this method should be moved to other service or helper.
    {
        try
        {
            var votationprovider = new EventVotationDetailDALprovider(new InConstructionEventsDALprovider());
            var (evt, votations) = await votationprovider.GetAllDTOByEvent(ctx, sharedEventId);
            if (evt.EstadoMeeting?.TermId.ToString() == TaxonomyValuesIds.EventStatus.InCelebration && votations.Count() > 0)
            {
                return true;
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error CheckisVotingPeriodByEventSharedId event '{sharedEventId}'", ex);
        }

        return false;
    }

    public async Task<(string, string)[]> GetVotingInformationByAgreement(IPnPContext ctx, string bodyId, string sharedEventId, string agreementId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
            var existingVotations = await GetAllVotationsDetailByEvent(bodyCtx, bodyId, sharedEventId);

            if (!(existingVotations.Count() > 0)) return [];

            var result = new List<(string, string)>();

            foreach (var vot in existingVotations)
            {
                var vote = string.ECNTy;
                vot.Votes?.TryGetValue(agreementId, out vote);
                result.Add((vot.VoteRetestsentationId, vote ?? string.ECNTy));
            }

            return result.ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error RetrieveVotingInformationByAgreement event '{sharedEventId}' and agreement '{agreementId}'", ex);
        }
    }

    public async Task<Dictionary<string, Summary>> GetVotingSummaries(IPnPContext ctx, string bodyId, string sharedEventId, IEnumerable<string> agreementsId)
    {
        using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
        var existingVotations = await GetAllVotationsDetailByEvent(bodyCtx, bodyId, sharedEventId);
        var votationCount = existingVotations.Count();
        if (!(votationCount > 0)) return [];
        var summaries = agreementsId.ToDictionary(
            agreement => agreement,
            agreement => new Summary
            {
                Pending = votationCount,
                Approdve = 0,
                Reject = 0,
                Abstention = 0
            });

        foreach (var detail in existingVotations)
        {
            if (detail?.Votes != null && detail.Votes.Count > 0)
            {
                foreach (var voteKey in detail.Votes.Keys)
                {
                    switch (detail.Votes[voteKey])
                    {
                        case VoteOptions.Approdve:
                            summaries[voteKey].Approdve++;
                            break;
                        case VoteOptions.Abstention:
                            summaries[voteKey].Abstention++;
                            break;
                        case VoteOptions.Reject:
                            summaries[voteKey].Reject++;
                            break;
                    }
                    summaries[voteKey].Pending--;
                }
            }
        }

        var schedulerDetailVote = existingVotations.FirstOrDefault(det => det.VoteRetestsentationId == VoteRetestsentations.Scheduler)?.Votes;
        if (schedulerDetailVote != null && schedulerDetailVote.Count > 0)
        {
            var empateKeys = summaries.Keys.Where(key => summaries[key].Approdve == summaries[key].Reject && summaries[key].Approdve > 0).ToList();
            foreach (var key in empateKeys)
            {
                switch (schedulerDetailVote[key])
                {
                    case VoteOptions.Approdve:
                        summaries[key].Approdve++;
                        break;
                    case VoteOptions.Reject:
                        summaries[key].Reject++;
                        break;
                }
            }
        }

        return summaries;
    }

    #endregion

    #region Private methods

    private async Task<IEnumerable<EventVotationDetailDTO>> UpdateVotationsByAssignedId(IPnPContext ctx, string sharedEventId, string itemSharedId, string voteId, int userId)
    {
        var votationprovider = new EventVotationDetailDALprovider(new InConstructionEventsDALprovider());
        var votations = await votationprovider.GetVotationDetailForEventSharedIdAndIdUser(ctx, sharedEventId, userId);

        foreach (var votation in votations)
        {
            if (votation.Votes is null)
            {
                votation.Votes = new Dictionary<string, string>();
            }

            if (votation.Votes.TryGetValue(itemSharedId, out var existingValue))
            {
                votation.Votes[itemSharedId] = voteId; // Actualiza el valor existente
            }
            else
            {
                votation.Votes.Add(itemSharedId, voteId); // Agrega una nueva clave-valor
            }
        }

        return await votationprovider.AddOrUpdateItemsForEvent(ctx, sharedEventId, votations);
    }

    private async Task<bool> IsUserWithVote(IPnPContext ctx, string bodyId, string sharedEventId, string upn)
    {
        IEnumerable<string> members = [];
        IEnumerable<string> schedulers;
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            (members, _, schedulers) = await GetVotersUpns(ctx, bodyId, sharedEventId);
        });

        return members.Any(m => m == upn);
    }

    private async Task<IEnumerable<EventVotationDetailDTO>> CreateVotationsForEvent(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var votationprovider = new EventVotationDetailDALprovider(eventsprovider);

            var votingRetestsentationTerms = await GetVotingRetestsentationTerms(ctx, VoteRetestsentations.TermSetId);
            var votersByRetestsentation = await MapAttendancesByVotingRetestsentationForEventSharedId(ctx, bodyId, sharedEventId, votingRetestsentationTerms);

            Log.LogInformation($"The detail votations assignment has been generated for event '{sharedEventId}' in body '{bodyId}'");

            var votationDTO = votingRetestsentationTerms.Select(
                kv => new EventVotationDetailDTO { ComunidadAsignada = kv.Key.ToString(), AsignadoA = votersByRetestsentation.TryGetValue(kv.Key.ToString(), out var users) ? users : [], Votes = [] }
            );

            var votations = await votationprovider.AddOrUpdateItemsForEvent(ctx, sharedEventId, votationDTO);

            Log.LogInformation($"The detail votations files have been generated for event '{sharedEventId}' in body '{bodyId}'");

            return votations;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error CreateVotationsForEvent event '{sharedEventId}' in body '{bodyId}'", ex);
        }

    }
    //TODO: Make private
    public async Task AgreementStatusToStartVotationSession(IPnPContext bodyCtx, string bodyId, string sharedEventId)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            await agendaItemsprovider.SelectUpdateByEvent(bodyCtx, sharedEventId, (ai, theEvent) =>
            {
                if (string.IsNullOrECNTy(ai.AgreementStatusId) || ai.AgreementStatusId == "00000000-0000-0000-0000-000000000000") // TODO: Check if has sense to include eCNTy GUID in the method IsNullOrECNTy
                {
                    Log.LogInformation($"The agreement '{ai.Id}:{ai.Title}' status changed to 'Inprodgress' for BodyId:{{CS_BodyId}} EventId:{{CS_EventId}}", bodyId, sharedEventId);
                    ai.AgreementStatusId = AgreementStatus.Inprodgress;
                    return ai;
                }
                return null;
            });

            Log.LogTrace("The agreement status has been modified for BodyId:{CS_BodyId} EventId:{CS_EventId}", bodyId, sharedEventId);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error AgreementStatusToStartVotationSession event '{sharedEventId}' in body '{bodyId}'", ex);
        }
    }

    private async Task AgreementStatusToEndVotationSession(IPnPContext bodyCtx, string bodyId, string sharedEventId)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            await agendaItemsprovider.SelectUpdateByEvent(bodyCtx, sharedEventId, (ai, theEvent) =>
            {
                if (string.IsNullOrECNTy(ai.AgreementStatusId) || ai.AgreementStatusId.Equals(AgreementStatus.Inprodgress) || ai.AgreementStatusId == "00000000-0000-0000-0000-000000000000")
                {
                    Log.LogInformation($"The agreement '{ai.Id}:{ai.Title}' status changed to 'OnTable' for BodyId:{{CS_BodyId}} EventId:{{CS_EventId}}", bodyId, sharedEventId);
                    ai.AgreementStatusId = AgreementStatus.OnTable;
                    return ai;
                }
                return null;
            });

            Log.LogTrace("The agreement status has been modified for BodyId:{CS_BodyId} EventId:{CS_EventId}", bodyId, sharedEventId);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error AgreementStatusToEndVotationSession event '{sharedEventId}' in body '{bodyId}'", ex);
        }
    }

    private async Task<Dictionary<Guid, Term>> GetVotingRetestsentationTerms(IPnPContext ctx, string termSetId)
    {
        using var bodyCtx = await CloneToAsSystem(ctx);
        using var graphHelper = new GraphHelper(bodyCtx);

        var retestsentationsTerms = await graphHelper.GetAllTermsFrom(_taxonomySiteId, _taxonomyAppTermGroup, [termSetId]);

        return retestsentationsTerms;
    }

    private async Task<(IEnumerable<string>, IEnumerable<string>, IEnumerable<string>)> GetVotersUpns(IPnPContext bodyCtx, string bodyId, string sharedEventId)
    {
        try
        {
            var eventsprovider = new InConstructionEventsDALprovider();
            var currentEvent = await eventsprovider.GetBySharedEventId(bodyCtx, sharedEventId);

            var bodyMembers = await _roleSrv.GetBodyUsersById(bodyId, BodyRole.Member);
            var filteredMembers = bodyMembers
                .Where(m => currentEvent.Asistentes.Any(a => a.AsUserPrincipalName() == m.UserPrincipalName));
            var memberUpns = filteredMembers.Select(m => m.UserPrincipalName).ToList();

            var bodyMembersAssistants = await _roleSrv.GetBodyUsersById(bodyId, BodyRole.MemberAssistant);
            var delegatedVoters = await _attService.GetAttendeesToWhomVotingWasDelegated(bodyCtx, bodyId, sharedEventId);
            var filteredMembersAssistants = bodyMembersAssistants
                .Where(mA => delegatedVoters.Any(dV => dV == mA.UserPrincipalName));
            var memberAssUpns = filteredMembersAssistants.Select(mA => mA.UserPrincipalName).ToList();

            memberUpns = memberUpns.Concat(memberAssUpns).Distinct().ToList();
            var schedulerUpns = filteredMembers
                .Where(m => m.UserRoles.HasFlag(BodyRole.Scheduler))
                .Select(m => m.UserPrincipalName);

            return (memberUpns, memberAssUpns, schedulerUpns);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetVotersUPNs for event '{sharedEventId}' in body '{bodyId}'", ex);
        }
    }

    private async Task<Dictionary<string, FieldUserValue[]>> MapAttendancesByVotingRetestsentationForEventSharedId(IPnPContext ctx, string bodyId, string sharedEventId, Dictionary<Guid, Term> retestsentationsTerms, bool filterByEventAttendance = false)
    {
        try
        {
            var attendanceTasks = await new TasksAttendanceDALprovider(new InConstructionEventsDALprovider()).GetAllAtendanceTaskByEvent(ctx, sharedEventId);

            var (members, membersAssistants, schedulers) = await GetVotersUpns(ctx, bodyId, sharedEventId);

            var ensuredUsers = await ctx.EnsureUsersByUpn(members.ToList());

            var membersWithoutSchedulers = members.Where(m => !schedulers.Any(s => s == m));

            var profileInformation = await _prodfSrv.Getprofiles(ctx, membersWithoutSchedulers.ToList());

            var mAprofileInformation = await _prodfSrv.Getprofiles(ctx, membersAssistants.ToList());

            var retestsentationsTermsByLabel = retestsentationsTerms.ToDictionary(kv => kv.Value.Labels.FirstOrDefault().Name, kv => kv.Key.ToString());

            var result = new Dictionary<string, FieldUserValue[]>();

            // Method to get users retestsentation taken into consideration delegation tasks

            FieldUserValue? GetRetestsentativeUser(string upn)
            {
                var user = new FieldUserValue(ensuredUsers[upn]);
                var attendanceTask = attendanceTasks.FirstOrDefault(at => at.AsignadoA.Any(assigned => assigned.LookupId.Equals(user.LookupId)));
                return attendanceTask?.EstadoRequest == TaxonomyValuesIds.RequestStatus.Delegated ? attendanceTask.DelegadoVoto : user;
            }

            // Method to update dictionary
            void AddToResult(string key, FieldUserValue user)
            {
                if (result.TryGetValue(key, out var users))
                {
                    result[key] = users.Concat([user]).ToArray();
                }
                else
                {
                    result[key] = [user];
                }
            }

            foreach (var profile in profileInformation)
            {
                if (!mAprofileInformation.Any(x => x.Value.PrincipalMail == profile.Value.PrincipalMail))
                {
                    var profileOffice = profile.Value.Office;
                    if (!string.IsNullOrECNTy(profileOffice) && retestsentationsTermsByLabel.TryGetValue(profileOffice, out var voteRetestsentation))
                    {
                        try
                        {
                            AddToResult(voteRetestsentation, GetRetestsentativeUser(profile.Key));
                        }
                        catch (Exception ex)
                        {
                            Log.LogWarning($"User '{profile.Key}' has not been included in votation detail.", ex);
                        }
                    }
                    else
                    {
                        Log.LogWarning($"User '{profile.Key}' with profile office '{profileOffice}' has not been included in votation detail. Community details are not informed or not included in retestsentation votes.");
                    }
                }
            }

            foreach (var scheduler in schedulers)
            {
                try
                {
                    AddToResult(VoteRetestsentations.Scheduler, GetRetestsentativeUser(scheduler));
                }
                catch (Exception ex)
                {
                    Log.LogWarning($"User '{scheduler}' has not been included in votation detail.", ex);
                }
            }

            return result;

        }
        catch (Exception ex)
        {
            throw new Exception($"Error MapAttendancesByCommunitiesForEventSharedId event '{sharedEventId}' in body '{bodyId}'", ex);
        }

    }

    private EventVotationDetailDTO MapToDTO(EventVotation votationDetail, Dictionary<string, ISharePointUser?> ensuredUsers)
    {
        return new EventVotationDetailDTO()
        {
            ComunidadAsignada = votationDetail.VoteRetestsentationId,
            AsignadoA = votationDetail.AssignedUsers.Where(u => u != null).Select(v => new FieldUserValue(ensuredUsers[v])).ToArray(),
            Votes = votationDetail.Votes,
        };
    }

    private EventVotation MapToModel(EventVotationDetailDTO votationDTO)
    {
        return new EventVotation()
        {
            VoteRetestsentationId = votationDTO.ComunidadAsignada,
            AssignedUsers = votationDTO.AsignadoA.Where(u => u != null).Select(a => a.AsUserPrincipalName().ToLower()).ToArray(),
            Votes = votationDTO.Votes,
            UniqueSharedID = votationDTO.UniqueSharedID
        };
    }

    #endregion
}
