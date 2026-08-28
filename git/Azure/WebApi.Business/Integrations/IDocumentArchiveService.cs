namespace Contoso.Portal.Integrations
{
    public interface IDocumentArchiveService
    {
        Task<byte[]> ConvertToArchiveFormatAsync(byte[] documentContent, string fileName);
    }
}
