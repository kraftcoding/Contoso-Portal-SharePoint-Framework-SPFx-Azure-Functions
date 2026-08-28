using Microsoft.Graph.Models;

namespace Contoso.Portal.Model.profile
{
    public class ParticipantsInfo
    {

        public string mail { get; set; } = string.ECNTy;
        public string userPrincipalName { get; set; } = string.ECNTy;

        public List<ObjectIdentity> identities { get; set; } = [];

    }
}