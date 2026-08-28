using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Model.Bodies;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.Bodies;

public class BodyRoleService(IMemoryCache memoryCache, ILogger<BodyRoleService> logger, M365AuthHelper auth, int cacheExpirationMinutes = 10) : ServiceBasePnP<BodyRoleService>(logger, auth)
{
    private readonly IMemoryCache _memoryCache = memoryCache;
    private readonly int _cacheExpirationMinutes = cacheExpirationMinutes;

    /// <summary>
    /// Checks if current user context is editor of body (Scheduler or SchedulerAssistant)
    /// </summary>
    /// <param name="ctx"></param>
    /// <param name="bodyId"></param>
    /// <returns></returns>
    public async Task<bool> CurrentUserIsEditorOfBodyOrSystem(IPnPContext ctx, string bodyId) => await ctx.IsElevatedOrSystem() || UserIsEditorOfBody(await GetBodiesFromUserContext(ctx), bodyId);
    public async Task<bool> CurrentUserIsMemberAssistant(IPnPContext ctx, string bodyId) => UserIsMemberAssistant(await GetBodiesFromUserContext(ctx), bodyId);
    public bool UserIsEditorOfBody(UserBodyRoles bodies) => bodies.UserRoles.HasFlag(BodyRole.Scheduler) || bodies.UserRoles.HasFlag(BodyRole.SchedulerAssistant);
    private bool UserIsEditorOfBody(IEnumerable<UserBodyRoles> allBodies, string bodyId) => allBodies.Where(ub => ub.Body.Id.Equals(bodyId, StringComparison.OrdinalIgnoreCase) && UserIsEditorOfBody(ub)).Any();
    public bool UserIsMemberAssistant(UserBodyRoles bodies) => bodies.UserRoles.HasFlag(BodyRole.MemberAssistant);
    private bool UserIsMemberAssistant(IEnumerable<UserBodyRoles> allBodies, string bodyId) => allBodies.Where(ub => ub.Body.Id.Equals(bodyId, StringComparison.OrdinalIgnoreCase) && UserIsMemberAssistant(ub)).Any();
    public async Task<bool> CurrentUserIsOCP(IPnPContext ctx)
    {
        var upn = await ctx.GetCurrentUserUpn();
        var isOCP = false;
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            using var graphHelper = new GraphHelper(ctxSystem);
            var groups = await graphHelper.GetUserGroups(upn);
            isOCP = groups.Any(group => group.DisplayName.Equals(DALConstants.AdminGroupNames.OCP, StringComparison.OrdinalIgnoreCase));
        });
        return isOCP;
    }

    public async Task<bool> UserIsAdmin(IPnPContext ctx, string upn)
    {
        var isAdmin = false;
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            var graphHelper = new GraphHelper(ctxSystem);
            var groups = await graphHelper.GetUserGroups(upn);
            isAdmin = groups.Any(group => group.DisplayName.Equals(DALConstants.AdminGroupNames.Admin, StringComparison.OrdinalIgnoreCase) || group.DisplayName.Equals(DALConstants.AdminGroupNames.Soporte, StringComparison.OrdinalIgnoreCase));
        });
        return isAdmin;
    }

    public async Task<bool> UserIsEXTERNALReader(IPnPContext ctx, string upn)
    {
        var isReader = false;
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            var graphHelper = new GraphHelper(ctxSystem);
            var groups = await graphHelper.GetUserGroups(upn);
            isReader = groups.Any(group => group.DisplayName.Equals(DALConstants.EXTERNALGroupNames.Reader, StringComparison.OrdinalIgnoreCase));
        });
        return isReader;
    }

    public async Task<IEnumerable<UserBodyRoles>> GetBodiesFromUserContext(IPnPContext anyUserCtx, BodyRole? userRoleFilter = null, BodyType? bodyTypeFilter = null)
    {
        return await GetBodiesForUser(await anyUserCtx.GetCurrentUserUpn(), userRoleFilter, bodyTypeFilter);
    }

    public async Task<IEnumerable<UserBodyRoles>> GetBodiesForUser(string userUPN, BodyRole? userRoleFilter = null, BodyType? bodyTypeFilter = null)
    {
        var groups = await GetUserGroupsNames(userUPN);

        return BodyPatternUtilities.ParseCollectionGroupNamesUsingNamePattern(groups)
            .Where(r =>
                (bodyTypeFilter == null || r.Type == bodyTypeFilter.Value) &&
                (userRoleFilter == null || r.UserRoles.HasFlag(userRoleFilter.Value))
            )
            .Select(r => new UserBodyRoles()
            {
                Body = new BodyBaseInfo() { Id = r.Id, RelativeUrl = r.RelativeUrl, Type = r.Type, TypeId = r.TypeId },
                UserPrincipalName = userUPN,
                UserRoles = r.UserRoles
            });
    }

    public async Task<IEnumerable<UserBodyRoles>> GetAllBodies(BodyRole? userRoleFilter = null, BodyType? bodyTypeFilter = null)
    {
        var groups = await GetAllBodiesIds();

        return BodyPatternUtilities.ParseCollectionGroupNamesUsingNamePattern(groups)
            .Where(r =>
                (bodyTypeFilter == null || r.Type == bodyTypeFilter.Value) &&
                (userRoleFilter == null || r.UserRoles.HasFlag(userRoleFilter.Value))
            )
            .Select(r => new UserBodyRoles()
            {
                Body = new BodyBaseInfo() { Id = r.Id, RelativeUrl = r.RelativeUrl, Type = r.Type, TypeId = r.TypeId },
                UserPrincipalName = string.ECNTy,
                UserRoles = r.UserRoles
            });
    }
    public async Task<UserBodyRoles?> GetRolesByBody(string userUPN, string bodyId)
    {
        var groups = await GetUserGroupsNames(userUPN);
        groups = groups.Where(x => x.StartsWith(bodyId));
        if (!groups.Any())
            return null;

        return BodyPatternUtilities.ParseCollectionGroupNamesUsingNamePattern(groups)
            .Select(r => new UserBodyRoles()
            {
                Body = new BodyBaseInfo() { Id = r.Id, RelativeUrl = r.RelativeUrl, Type = r.Type, TypeId = r.TypeId },
                UserPrincipalName = userUPN,
                UserRoles = r.UserRoles
            }).FirstOrDefault();
    }

    public async Task<IEnumerable<UserBodyRoles>> GetBodyUsersById(string bodyId, BodyRole? userRoleFilter = null)
    {
        var groupsAndUsers = await GetGroupsAndMembersOfBodyById(bodyId);
        var result = new Dictionary<string, BodyRole>();

        BodyBaseInfo? bodyInformation = null;
        var bodyInfo = BodyPatternUtilities.ParseUsingNamePattern(bodyId);
        if (bodyInfo is not null)
        {
            var (id, url, type, typeId, _, _, _, _) = bodyInfo.Value;
            bodyInformation = new BodyBaseInfo()
            {
                Id = id,
                RelativeUrl = url,
                Type = type,
                TypeId = typeId
            };

            foreach (var (groupName, members) in groupsAndUsers)
            {
                var groupInfo = BodyPatternUtilities.ParseUsingNamePattern(groupName);
                if (groupInfo is null)
                    continue;

                var (_, _, _, _, role, _, _, _) = groupInfo.Value;

                var membersOfGroup = members
                .Select(m => m.ToLower());

                foreach (var userInGroup in membersOfGroup)
                    if (result.ContainsKey(userInGroup))
                        result[userInGroup] |= role;
                    else
                        result.Add(userInGroup, role);
            }
        }

        if (bodyInformation is null || result.Count == 0)
            return [];

        return result
            .Where(r => userRoleFilter == null || r.Value.HasFlag(userRoleFilter.Value))
            .Select(r => new UserBodyRoles() { Body = bodyInformation, UserPrincipalName = r.Key, UserRoles = r.Value });
    }

    public async Task<IEnumerable<string>> GetBodyGroupsFromUser(string bodyId, string userUPN)
    {
        var groupsAndUsers = await GetGroupsAndMembersOfBodyById(bodyId);
        var result = new List<string>();

        foreach (var (GroupName, Members) in groupsAndUsers)
            if (Members.Where(m => m.Equals(userUPN, StringComparison.OrdinalIgnoreCase)).Any())
                result.Add(GroupName);

        return result;
    }

    public async Task<IEnumerable<string>> GetAllBodiesIds()
    {
        return await GetAllBodiesIdsFromMatchingGroups();
    }

    internal async Task<IEnumerable<string>> GetUserGroupsNames(string userUPN)
    {
        return await _memoryCache.GetOrCreateAsync(userUPN, async (entry) =>
        {
            entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(_cacheExpirationMinutes);

            using var rootCtx = await CreatePnPContextAsSystem();
            using var graphHelper = new GraphHelper(rootCtx);

            var userGroups = await graphHelper.GetUserGroups(userUPN);
            return userGroups?.Select(g => g.DisplayName!) ?? [];
        }) ?? [];
    }

    internal async Task<IEnumerable<string>> GetAllBodiesIdsFromMatchingGroups()
    {
        return await _memoryCache.GetOrCreateAsync("allBodiesIds", async (entry) =>
        {
            using var rootCtx = await CreatePnPContextAsSystem();
            using var graphHelper = new GraphHelper(rootCtx);

            var groups = await graphHelper.GetGroups(["displayName"]);
            var parsedResults = BodyPatternUtilities.ParseCollectionGroupNamesUsingNamePattern(groups.Select(g => g.DisplayName!));

            return parsedResults.Select(r => r.Id).Distinct();
        }) ?? [];
    }

    internal async Task<IEnumerable<(string GroupName, IEnumerable<string> Members)>> GetGroupsAndMembersOfBodyById(string bodyId)
    {
        return await _memoryCache.GetOrCreateAsync(bodyId, async (entry) =>
        {
            entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(_cacheExpirationMinutes);

            using var rootCtx = await CreatePnPContextAsSystem();
            using var graphHelper = new GraphHelper(rootCtx);

            var graphGroups = await graphHelper.GetGroupsMembersWhenNameStartsWith(bodyId);

            return graphGroups.Select(g => (
                GroupName: g.Group.DisplayName!,
                MembersUpns: g.Members.Where(m => !string.IsNullOrECNTy(m.UserPrincipalName)) // #ext# users are having userPrincipalName eCNTy (graph batch bug?)
                    .Select(m => m.UserPrincipalName!)));
        }) ?? [];
    }

    public async Task<IEnumerable<string>> GetMembersOfBodyByIdAndRole(string bodyId, BodyRole[] roles)
    {
        using var rootCtx = await CreatePnPContextAsSystem();
        using var graphHelper = new GraphHelper(rootCtx);

        var groupnames = roles.Select(role => BodyPatternUtilities.RoleToGroupName(bodyId, role)).ToList();
        var graphGroups = await graphHelper.GetGroupsMembersWhenNameStartsWith(bodyId);

        var filteredGroups = graphGroups.Where(g => groupnames.Any(name => g.Group.DisplayName.Equals(name, StringComparison.OrdinalIgnoreCase)));

        return filteredGroups
             .SelectMany(g => g.Members.Where(m => !string.IsNullOrECNTy(m.UserPrincipalName))
             .Select(m => m.UserPrincipalName!))
             .ToHashSet()
             .ToArray();

    }

    public async Task<Source> GetDefaultEventSourceByBodyRole(IPnPContext bodyCtx, string bodyId, bool fromArchive = false)
    {
        return fromArchive ? Source.Archived : (await CurrentUserIsEditorOfBodyOrSystem(bodyCtx, bodyId) ? Source.InConstruction : Source.Published);
    }

    // TODO: mover a un helper o similar. No tiene que ver con los roles de los cuerpos
    public string GetNotificationUrlByEvent(IPnPContext ctx, string sharedEventId, string bodyId, bool isDocs = false)
    {
        var uriBuilder = new UriBuilder(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(BodyPatternUtilities.BodyIdToSiteUrl(bodyId))));
        uriBuilder.Fragment += string.Format("/{0}", sharedEventId) + (isDocs ? "/docs" : "");
        return uriBuilder.Uri.ToString();
    }
}