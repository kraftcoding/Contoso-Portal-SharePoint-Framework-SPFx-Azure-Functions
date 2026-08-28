namespace Contoso.Portal.Integrations
{
    public interface INotificationSmsService
    {
        Task SendSmsAsync(IEnumerable<string> phoneNumbers, string message);
    }
}
