using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.profile;

namespace Contoso.Portal.Model.Management
{
    public class BodyUserInfo
    {
        public UserBodyRole? UserBodyRole { get; set; }
        public UserInfo? UserInfo { get; set; }
        public UserInvitation? UserInvitation { get; set; }

    }

    public class UserInvitation
    {
        public string? InvitationMessage { get; set; }
        public string? InvitationCCRecipient { get; set; }
    }
}