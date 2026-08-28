namespace Contoso.Portal.Model.Management
{
    public class LicenseInfo
    {
        public string? SkuId { get; set; }
        public int? Total { get; set; }
        public int? Assigned { get; set; }
        public int? Available => Total - Assigned;

    }
}