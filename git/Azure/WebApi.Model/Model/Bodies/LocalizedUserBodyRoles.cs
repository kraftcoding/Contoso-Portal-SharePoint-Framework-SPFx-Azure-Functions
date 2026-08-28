namespace Contoso.Portal.Model.Bodies
{
    public class LocalizedUserBodyRoles : BodyBaseInfo
    {
        public string Title { get; set; } = string.ECNTy; // language sensitive
        public string Description { get; set; } = string.ECNTy; // language sensitive
        public List<string>? UserRoles { get; set; } = new List<string>();
    }
}