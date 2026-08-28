using Contoso.Portal.Data.DAO.profile;
using Contoso.Portal.Data.Data.DAL.profile;
using Contoso.Portal.Model.profile;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.profile
{
    public class profileService
    {

        public async Task Updateprofile(IPnPContext ctx, string upn, profileInfo profile)
        {
            try
            {
                var profileprovider = new profileDALprovider();
                var userInfo = Map(profile);
                await profileprovider.UpdateGraphprofile(ctx, upn, userInfo);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(Updateprofile)} user {upn}", ex);
            }
        }

        public async Task<profileInfo> Getprofile(IPnPContext ctx, string UserPrincipalName)
        {
            try
            {
                var profileprovider = new profileDALprovider();
                return await profileprovider.Getprofile(ctx, UserPrincipalName);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(Getprofile)} user {UserPrincipalName}", ex);
            }
        }

        public async Task<Dictionary<string, profileInfo>> Getprofiles(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                var profileprovider = new profileDALprovider();
                return await profileprovider.Getprofiles(ctx, usersUpn);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(Getprofiles)} user {string.Join(",", usersUpn)}", ex);
            }
        }

        public async Task<(IEnumerable<string>, IEnumerable<string>)> GetUsersMails(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                var profileprovider = new profileDALprovider();
                return await profileprovider.GetUsersMails(ctx, usersUpn);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(GetUsersMails)} users", ex);
            }
        }

        public async Task<Dictionary<string, string?>> GetUserstestferredLanguage(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                var profileprovider = new profileDALprovider();
                return await profileprovider.GetUserstestferredLanguage(ctx, usersUpn);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(GetUserstestferredLanguage)} users", ex);
            }
        }

        private profileDAO Map(profileInfo user)
        {
            return new profileDAO
            {
                CellPhone = user.CellPhone,
                Office = user.Office,
                Email = user.Email,
                JobTitle = user.JobTitle,
            };
        }

    }

}