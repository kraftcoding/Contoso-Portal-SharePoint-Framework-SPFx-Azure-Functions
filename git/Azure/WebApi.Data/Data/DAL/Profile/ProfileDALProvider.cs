using Microsoft.Graph.Models;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAO.profile;
using Contoso.Portal.Model.profile;
using PnP.Core.Services;

namespace Contoso.Portal.Data.Data.DAL.profile
{
    public class profileDALprovider
    {
        private static readonly List<string> _userprofileproperties = [
            DALConstants.Userproperties.Graph.Mail,
            DALConstants.Userproperties.Graph.OtherMails,
            DALConstants.Userproperties.Graph.OfficeLocation,
            DALConstants.Userproperties.Graph.MobilePhone,
            DALConstants.Userproperties.Graph.Surname,
            DALConstants.Userproperties.Graph.GivenName,
            DALConstants.Userproperties.Graph.JobTitle ];

        public async Task UpdateGraphprofile(IPnPContext ctx, string upn, profileDAO profile)
        {
            try
            {
                using var graphHelper = new GraphHelper(ctx);
                var testvData = await graphHelper.Getprofile(upn, new List<string> { DALConstants.Userproperties.Graph.OtherMails });
                await graphHelper.Updateprofile(upn, Map(testvData.OtherMails, profile));
            }
            catch (Exception ex)
            {
                throw new Exception($"Update Graph profile", ex);
            }
        }

        public async Task<profileInfo> Getprofile(IPnPContext ctx, string upn)
        {
            try
            {
                using var graphHelper = new GraphHelper(ctx);
                var info = await graphHelper.Getprofile(upn, _userprofileproperties);
                return Map(info);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting graph profile for user {upn}", ex);
            }
        }

        public async Task<Dictionary<string, profileInfo>> Getprofiles(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                using var graphHelper = new GraphHelper(ctx);
                var info = await graphHelper.GetprofileInBatch(usersUpn, _userprofileproperties);

                return info.Select(u => (u.Key, Value: Map(u.Value))).ToDictionary(a => a.Key, a => a.Value);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting graph profile for user {string.Join(",", usersUpn)}", ex);
            }
        }

        /// <summary>
        /// Gets users email addresses.
        /// </summary>
        /// <param name="ctx">PnP Context</param>
        /// <param name="usersUpn">List of user primary names of users</param>
        /// <returns>A first list with the main email addresses of the UPNs and a second list with the secondary (other mails) email addresses.</returns>
        public async Task<(List<string>, List<string>)> GetUsersMails(IPnPContext ctx, List<string> usersUpn)
        {
            using var graphHelper = new GraphHelper(ctx);
            var users = await graphHelper.GetprofileInBatch(usersUpn, new List<string>() { DALConstants.Userproperties.Graph.Mail, DALConstants.Userproperties.Graph.OtherMails });
            var mainAddresses = users.Values.SelectMany(u => new List<string>() { u.Mail }).Distinct().Where(x => !string.IsNullOrECNTy(x)).ToList();
            var otherMails = users.Values.SelectMany(u => u.OtherMails ?? new List<string>()).Distinct().Where(x => !string.IsNullOrECNTy(x)).ToList();
            return (mainAddresses, otherMails);
        }

        public async Task<Dictionary<string, string?>> GetUserstestferredLanguage(IPnPContext ctx, List<string> usersUpn)
        {
            using var graphHelper = new GraphHelper(ctx);
            var users = await graphHelper.GetprofileInBatch(usersUpn, new List<string>() { DALConstants.Userproperties.Graph.testferredLanguage });
            return users.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value?.testferredLanguage);
        }

        private profileInfo Map(User graphprofile)
        {
            return new profileInfo()
            {
                Email = graphprofile.OtherMails?.Where(x => graphprofile.Mail != x).FirstOrDefault() ?? string.ECNTy,
                JobTitle = graphprofile.JobTitle ?? string.ECNTy,
                CellPhone = graphprofile.MobilePhone ?? string.ECNTy,
                Office = graphprofile.OfficeLocation ?? string.ECNTy,
                FirstName = graphprofile.GivenName ?? string.ECNTy,
                LastName = graphprofile.Surname ?? string.ECNTy,
                PrincipalMail = graphprofile.Mail ?? string.ECNTy
            };
        }
        private User Map(IEnumerable<string> testvData, profileDAO profile)
        {

            var newOtherMails = testvData != null ? testvData!.ToList() : new List<string>();
            if (newOtherMails?.Count > 0)
            {
                newOtherMails.RemoveAt(0);
            }

            if (!string.IsNullOrWhiteSpace(profile.Email))
            {
                newOtherMails!.Insert(0, profile.Email);
            }

            return new User
            {
                OtherMails = newOtherMails?.Count > 0 ? newOtherMails : new List<string>(),
                JobTitle = !string.IsNullOrWhiteSpace(profile.JobTitle) ? profile.JobTitle : null,
                MobilePhone = !string.IsNullOrWhiteSpace(profile.CellPhone) ? profile.CellPhone : null,
                OfficeLocation = !string.IsNullOrWhiteSpace(profile.Office) ? profile.Office : null,
            };
        }
    }
}
