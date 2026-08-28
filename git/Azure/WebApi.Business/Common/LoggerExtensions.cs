using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Common;

public static class ContosoScopeFormatterExtension
{
    public static IDisposable? BeginContosoScope(this ILogger logger, string message, string? bodyId = null, string? eventId = null, string? upn = null)
    {
        var scopeMessage = $"[SCOPE]: {message}: BodyId:{{CS_BodyId}} EventId:{{CS_EventId}} Upn:{{CS_UPN}}";
        // logger.LogTrace($"[SCOPE]: {message}"); // in AI BeginScope message is not logged but data will be added to custom dimensions
        return logger.BeginScope(scopeMessage, [bodyId ?? "N/A", eventId ?? "N/A", upn ?? "N/A"]);
    }
}