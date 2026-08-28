using Microsoft.Graph.Models;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Model.profile;
using PnP.Core.Services;

namespace Contoso.Portal.Data.Data.DAL.Helpers
{
    public static class MeetingParticipantsHelper
    {
        private static readonly List<string> _userTeamsprofileproperties = [
            DALConstants.Userproperties.ParticipantsInfo.Mail,
            DALConstants.Userproperties.ParticipantsInfo.UPN,
            DALConstants.Userproperties.ParticipantsInfo.Identities
        ];

        public static async Task<Dictionary<string, IEnumerable<string>>> GetUpnToMeetingParticipantUpn(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                var dictionary = await GetParticipantsInfo(ctx, usersUpn);
                return dictionary.ToDictionary(a => a.Key, a => GetParticipantUpn(a.Value));
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting UPN to Meeting particpants UPN. Exception: {ex}");
            }
        }

        private static async Task<Dictionary<string, ParticipantsInfo>> GetParticipantsInfo(IPnPContext ctx, List<string> usersUpn)
        {
            try
            {
                using var graphHelper = new GraphHelper(ctx);
                var info = await graphHelper.GetprofileInBatch(usersUpn, _userTeamsprofileproperties);

                return info.Select(u => (u.Key, Value: Map(u.Value))).ToDictionary(a => a.Key, a => a.Value);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error getting graph participants info for user {string.Join(",", usersUpn)}", ex);
            }
        }

        private static IEnumerable<string> GetParticipantUpn(ParticipantsInfo info)
        {
            if (info.mail is null || info.userPrincipalName is null || info.identities is null || info.identities?.Count == 0)
                throw new Exception($"Error getting participants info. Some mandatory field is missing. Userinfo: {info}");

            // add copera Upn and check for external identites (other tenant / microsoft account)
            var upns = new List<string>() { info.userPrincipalName };

            var userFederatedIdentities = info.identities?.Where(i => i.SignInType == DALConstants.Userproperties.ParticipantsInfo.Options.Federated) ?? [];
            foreach (var identity in userFederatedIdentities)
            {
                if (identity.Issuer == DALConstants.Userproperties.ParticipantsInfo.Options.ExternalAzure)
                {
                    upns.Add(info.mail);
                    break;
                }
            }

            return upns;
        }

        private static ParticipantsInfo Map(User graphprofile)
        {
            return new ParticipantsInfo()
            {
                mail = graphprofile.Mail,
                userPrincipalName = graphprofile.UserPrincipalName,
                identities = graphprofile.Identities
            };
        }
    }
}
