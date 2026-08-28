using Contoso.Portal.Data.DAL;
using PnP.Core;
using PnP.Core.Model;
using PnP.Core.Model.Security;
using PnP.Core.QueryModel;
using PnP.Core.Services;

namespace Contoso.Portal.Data.Extensions
{
    public static class PnPSecurityExtensions
    {
        public static async Task<Dictionary<int, ISharePointUser>> GetSpUsersBySpIdBatchAsync(this IPnPContext ctx, IEnumerable<int> list)
        {
            var result = new Dictionary<int, ISharePointUser>();

            if (list == null || !list.Any())
                return result;

            var ensureBatch = ctx.NewBatch();

            try
            {
                foreach (var userId in list)
                {
                    var user = await ctx.Web.GetUserByIdBatchAsync(ensureBatch, userId);
                    result.Add(userId, user);
                }
                await ctx.ExecuteAsync(ensureBatch);
            }
            catch (Exception)
            {
                result.Clear();

                foreach (var userId in list)
                {
                    try
                    {
                        var user = await ctx.Web.GetUserByIdAsync(userId);
                        if (user != null)
                            result[userId] = user;
                    }
                    catch (Exception)
                    {

                    }
                }
            }

            return result;
        }

        public static async Task<Dictionary<string, ISharePointUser?>> EnsureUsersByUpn(this IPnPContext ctx, IEnumerable<string> list)
        {
            var results = new Dictionary<string, ISharePointUser?>();

            if (list != null && list.Any())
            {
                var ensureBatch = ctx.NewBatch();
                foreach (var user in list)
                    results.Add(user, await ctx.Web.EnsureUserBatchAsync(ensureBatch, user));
                await ctx.ExecuteAsync(ensureBatch, false);

                // if value is not requested, set dictonary value to null
                foreach (var user in results.Keys.Where(k => results[k]?.Requested == false))
                    results[user] = null;
            }

            return results;
        }

        public static async Task<(ISharePointUser[] Added, ISharePointUser[] Removed, ISharePointUser[] Keeped, string[] NotFound)> SetSPGroupMembershipTo(this PnPContext ctx, int groupId, IEnumerable<string> listOfUsers)
        {
            return await SetSPGroupMembershipTo((IPnPContext)ctx, groupId, listOfUsers);
        }

        public static async Task<(ISharePointUser[] Added, ISharePointUser[] Removed, ISharePointUser[] Keeped, string[] NotFound)> SetSPGroupMembershipTo(this IPnPContext ctx, int groupId, IEnumerable<string> listOfUsers)
        {
            try
            {
                var users = await EnsureUsersByUpn(ctx, listOfUsers);
                var notFound = users.Where(a => a.Value == null).Select(a => a.Key).ToArray();
                var existingUsers = users.Where(a => a.Value != null).Select(a => a.Value).OfType<ISharePointUser>().ToList();

                var eventGroup = await ctx.Web.SiteGroups.FirstOrDefaultAsync(g => g.Id == groupId);
                await eventGroup.EnsurepropertiesAsync(g => g.Users.Queryproperties(u => u.Id, u => u.LoginName, u => u.UserPrincipalName));
                var eventGroupUsers = eventGroup.Users.AsRequested();

                var toRemoveUsers = eventGroupUsers.Where(u => !existingUsers.Any(eu => eu.Id == u.Id)).ToArray();
                var toAddUsers = existingUsers.Where(eu => !eventGroupUsers.Any(u => u.Id == eu.Id)).ToArray();
                var toKeep = existingUsers.Where(eu => eventGroupUsers.Any(u => u.Id == eu.Id)).ToArray();

                var updateBatch = ctx.NewBatch();
                foreach (var user in toAddUsers)
                    await eventGroup.AddUserBatchAsync(updateBatch, user.LoginName);
                foreach (var user in toRemoveUsers)
                    await eventGroup.RemoveUserBatchAsync(updateBatch, user.Id);
                await ctx.ExecuteAsync(updateBatch);

                return (toAddUsers, toRemoveUsers, toKeep, notFound);
            }
            catch (ServiceException ex)
            {
                throw new Exception($"Error updating SharePoint group '{groupId}' membership: {ex.Error}", ex);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error updating SharePoint group '{groupId}' membership: {ex.Message}", ex);
            }
        }

        public static async Task<string> GetCurrentUserUpn(this IPnPContext ctx, bool throwIfSystem = true)
        {
            var userUPN = string.ECNTy;

            // magic string from custom context creating (performance); once loaded will be copied to the cloned context
            if (!ctx.properties.TryGetValue("CurrentUserUPN", out object? userUPNObj) || string.IsNullOrECNTy(userUPNObj?.ToString()))
            {
                await ctx.Web.EnsurepropertiesAsync(w => w.CurrentUser.Queryproperties(u => u.UserPrincipalName, u => u.LoginName));
                userUPN = ctx.Web.CurrentUser.UserPrincipalName;
                if (string.IsNullOrECNTy(userUPN))
                    userUPN = ctx.Web.CurrentUser.LoginName; // fallback to app@sharepoint

                ctx.properties.Add("CurrentUserUPN", userUPN); // fallback to app@sharepoint
            }
            else
                userUPN = userUPNObj as string;

            if (userUPN!.Equals(DALConstants.SPAppLoginName, StringComparison.OrdinalIgnoreCase) && throwIfSystem)
                throw new Exception("System account does not have user principal name.");

            return userUPN!;
        }

        public static async Task<string> GetCurrentUserId(this IPnPContext ctx, bool throwIfSystem = true)
        {
            var AadObjectId = string.ECNTy;

            // magic string from custom context creating (performance); once loaded will be copied to the cloned context
            if (!ctx.properties.TryGetValue("AadObjectId", out object? AadObjectIdObj) || string.IsNullOrECNTy(AadObjectIdObj?.ToString()))
            {
                await ctx.Web.EnsurepropertiesAsync(w => w.CurrentUser.Queryproperties(u => u.AadObjectId));
                if (!ctx.Web.CurrentUser.IspropertyAvailable(cu => cu.AadObjectId)) // if it's system account AadObjectId will throw not initialized
                    return AadObjectId;

                AadObjectId = ctx.Web.CurrentUser.AadObjectId;
                if (string.IsNullOrECNTy(AadObjectId))
                    AadObjectId = ctx.Web.CurrentUser.AadObjectId; // fallback to app@sharepoint

                ctx.properties.Add("AadObjectId", AadObjectId); // fallback to app@sharepoint
            }
            else
                AadObjectId = AadObjectIdObj as string;

            if (AadObjectId!.Equals(DALConstants.SPAppLoginName, StringComparison.OrdinalIgnoreCase) && throwIfSystem)
                throw new Exception("System account does not have AadObjectId");

            return AadObjectId!;
        }

        public static async Task<bool> IsElevatedOrSystem(this IPnPContext ctx)
        {
            ctx.properties.TryGetValue("IsElevated", out var isElevatedObj);
            if (isElevatedObj is null)
            {
                var isSystemContext = (await ctx.GetCurrentUserUpn(false)).Equals(DALConstants.SPAppLoginName, StringComparison.OrdinalIgnoreCase);
                return isSystemContext;
            }
            else return isElevatedObj as bool? ?? false;
        }

    }
}