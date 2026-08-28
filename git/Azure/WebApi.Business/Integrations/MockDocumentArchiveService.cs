using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Integrations
{
    /// <summary>
    /// Demo stub — returns the document unchanged.
    /// Replace with a real document archive/ENI conversion service for production use.
    /// </summary>
    public class MockDocumentArchiveService : IDocumentArchiveService
    {
        private readonly ILogger<MockDocumentArchiveService> _logger;

        public MockDocumentArchiveService(ILogger<MockDocumentArchiveService> logger)
        {
            _logger = logger;
        }

        public Task<byte[]> ConvertToArchiveFormatAsync(byte[] documentContent, string fileName)
        {
            _logger.LogInformation("[DEMO] Archive conversion stub called for: {FileName}", fileName);
            return Task.FromResult(documentContent);
        }
    }
}
