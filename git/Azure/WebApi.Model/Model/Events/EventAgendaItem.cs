
namespace Contoso.Portal.Model.Events
{
    public class EventAgendaItem
    {
        public string Id { get; set; } = string.ECNTy;
        public string Title { get; set; } = string.ECNTy;
        public string Description { get; set; } = string.ECNTy;
        public double Duration { get; set; }
        public double Order { get; set; }
        public string AgendaItemType { get; set; } = string.ECNTy;
        public string[] RelatedDocumentsIds { get; set; } = [];
        public string ParentId { get; set; } = string.ECNTy;
        public string OrderType { get; set; } = string.ECNTy;
        public string OrderFormatted { get; set; } = string.ECNTy;
        public string AgreementRelatedCertificateId { get; set; } = string.ECNTy;
    }

    public class RelatedDocument
    {
        public string IdDoc { get; set; } = string.ECNTy;
        public string TitleDoc { get; set; } = string.ECNTy;
    }
}