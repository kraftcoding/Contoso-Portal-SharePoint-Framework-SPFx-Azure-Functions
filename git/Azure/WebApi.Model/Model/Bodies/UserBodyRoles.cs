namespace Contoso.Portal.Model.Bodies
{
    public class UserBodyRoles
    {
        public string UserPrincipalName { get; set; }
        public BodyBaseInfo Body { get; set; }
        public BodyRole UserRoles { get; set; }
    }
}