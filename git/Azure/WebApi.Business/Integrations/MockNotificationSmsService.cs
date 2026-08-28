using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Integrations
{
    /// <summary>
    /// Demo stub — logs SMS messages instead of sending them.
    /// Replace with a real SMS provider implementation for production use.
    /// </summary>
    public class MockNotificationSmsService : INotificationSmsService
    {
        private readonly ILogger<MockNotificationSmsService> _logger;

        public MockNotificationSmsService(ILogger<MockNotificationSmsService> logger)
        {
            _logger = logger;
        }

        public Task SendSmsAsync(IEnumerable<string> phoneNumbers, string message)
        {
            _logger.LogInformation("[DEMO] SMS to {Count} recipient(s): {Message}",
                phoneNumbers.Count(), message);
            return Task.CompletedTask;
        }
    }
}
