using System.Security.Cryptography.X509Certificates;

namespace Contoso.Portal.Runtime
{
    public class Settings
    {
        public string? TenantId { get; set; }
        public string? ClientId { get; set; }
        public string? StorageAccountName { get; set; }

        public StoreName CertificateStoreName { get; set; }
        public StoreLocation CertificateStoreLocation { get; set; }
        public string? CertificateThumbprint { get; set; }
        public string? RootSiteUrl { get; set; }
        public string? DefaultLocale { get; set; }
        public string? GraphMailboxUserId { get; set; }

        public string? TaxonomyRootSiteId { get; set; }
        public string? TaxonomyAppTermGroup { get; set; }
        public string? TaxonomyDepartmentsSetId { get; set; }

        public bool? VotingEnabled { get; set; }

        public bool IsProductionEnvironment { get; set; } = true;

        public bool Publishing_ReplicateSensitivityEnabled { get; set; } = false;
        public bool Publishing_PdfConversionEnabled { get; set; } = true;
    }
}
