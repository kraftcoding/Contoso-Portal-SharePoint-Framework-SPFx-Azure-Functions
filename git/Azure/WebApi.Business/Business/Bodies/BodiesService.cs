using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Model.Bodies;
using PnP.Core.Services;
using User = Contoso.Portal.Model.Bodies.User;

namespace Contoso.Portal.Domains.Bodies
{
    public class BodiesService : ServiceBasePnP<BodiesService>
    {
        private readonly BodyRoleService _roleSrv;

        public BodiesService(BodyRoleService roleSrv, ILogger<BodiesService> logger, M365AuthHelper auth) : base(logger, auth)
        {
            _roleSrv = roleSrv;
        }

        public async Task<IEnumerable<LocalizedUserBodyRoles>> GetUserInitialBodiesInformation(IPnPContext ctx)
        {
            try
            {
                var userBodies = await _roleSrv.GetBodiesFromUserContext(ctx);

                // Define the roles to filter out
                var rolesToFilterOut = new[] { BodyRole.Member, BodyRole.Scheduler, BodyRole.SchedulerAssistant, BodyRole.MemberAssistant };
                // Filter out bodies where the user's role is in the roles to filter out
                userBodies = userBodies.Where(ub => rolesToFilterOut.Any(role => ub.UserRoles.HasFlag(role)));
                List<LocalizedUserBodyRoles> localizedUserBodyRoles = new List<LocalizedUserBodyRoles>();
                //TODO Encontrar una mejor forma para iterar sobre el objeto y mejorar el codigo
                await Task.WhenAll(
                    userBodies.Select(async ub =>
                {
                    try
                    {
                        // TODO: review where to store the body title and description. This IS SLOW.
                        using var bodyCtx = await ctx.CloneAsync(new Uri(ctx.Uri.GetLeftPart(UriPartial.Authority).UriCombine(ub.Body.RelativeUrl)));
                        await bodyCtx.Web.EnsurepropertiesAsync(w => w.Description, w => w.Title);
                        localizedUserBodyRoles.Add(new LocalizedUserBodyRoles()
                        {
                            Id = ub.Body.Id,
                            Title = bodyCtx.Web.Title,
                            Description = bodyCtx.Web.Description,
                            RelativeUrl = ub.Body.RelativeUrl,
                            Type = ub.Body.Type,
                            TypeId = ub.Body.TypeId,
                            UserRoles = Enum.GetValues(typeof(BodyRole))
                                .Cast<BodyRole>()
                                .Where(r => ub.UserRoles.HasFlag(r)).Select(r => BodyPatternUtilities.RoleToString(r)).ToList()
                        }
                        );
                    }
                    catch (Exception ex)
                    {
                        Log.LogWarning($"Error getting body '{ub.Body.RelativeUrl}, '{ex}''");
                    }
                    return;
                }));
                return localizedUserBodyRoles;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error bodies from context", ex);
            }
        }

        public async Task<IEnumerable<User>> GetBodyUsersByRole(IPnPContext ctx, string bodyId, BodyRole? role)
        {
            try
            {
                var members = (await _roleSrv.GetBodyUsersById(bodyId, role))
                    .Select(ur => new User
                    {
                        UserPrincipalName = ur.UserPrincipalName,
                        UserRoles = Enum.GetValues(typeof(BodyRole))
                            .Cast<BodyRole>()
                            .Where(r => ur.UserRoles.HasFlag(r)).Select(r => BodyPatternUtilities.RoleToString(r)).ToList()
                    });

                return members;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting users from {bodyId}", ex);
            }
        }
    }
}