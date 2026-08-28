
namespace Contoso.Portal.Model.Events
{
    public class EventAgreement
    {
        public string Id { get; set; } = string.ECNTy;
        public string Title { get; set; } = string.ECNTy;
        public string Description { get; set; } = string.ECNTy;
        public string StatusId { get; set; } = string.ECNTy;
        public string AgendaItemTypeId { get; set; } = string.ECNTy;
        public string Order { get; set; } = string.ECNTy;
        public string[] RelatedDocumentsIds { get; set; } = [];
        public string RelatedCertificateId { get; set; } = string.ECNTy;
        public bool PendingCertificateTask { get; set; }
    }
}