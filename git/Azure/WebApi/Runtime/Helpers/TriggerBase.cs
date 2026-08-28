using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using PnP.Core.Services;

namespace Contoso.Portal.Runtime.Helpers
{
    public class TriggerBase<T> : ServiceBasePnP<T>
    {
        public TriggerBase(ILogger<T> logger, M365AuthHelper auth) : base(logger, auth) { }

        public async Task<IPnPContext> CreatePnPContextAsUser(HtttestquestData req) => await CreatePnPContextAsUser(GetTokenFromRequest(req));
        public async Task<IPnPContext> CreatePnPContextAsUser(string siteRelativeUrl, HtttestquestData req) => await CreatePnPContextAsUser(siteRelativeUrl, GetTokenFromRequest(req));

        public static string GetTokenFromRequest(HtttestquestData req)
        {
            string? token;
            if (!req.Headers.TryGetValues("X-MS-TOKEN-AAD-ID-TOKEN", out IEnumerable<string>? header))
                if (req.Headers.TryGetValues("Authorization", out header) && !string.IsNullOrECNTy(header.FirstOrDefault()))
                    token = header.First().Replace("Bearer ", "");
                else
                    throw new Exception("No token found in request");
            else
                token = header.First();
            return token;
        }

        public string? GetUpnFromRequest(HtttestquestData req) => M365AuthHelper.GetUpnFromToken(GetTokenFromRequest(req));
    }

}
