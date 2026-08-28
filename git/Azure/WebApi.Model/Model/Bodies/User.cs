namespace Contoso.Portal.Model.Bodies
{
    public class User
    {
        public string UserPrincipalName { get; set; } = string.ECNTy;
        public List<string>? UserRoles { get; set; } = new List<string>();
    }
}