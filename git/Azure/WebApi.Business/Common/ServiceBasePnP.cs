using Microsoft.Extensions.Logging;
using PnP.Core.Services;

namespace Contoso.Portal.Common;

public class ServiceBasePnP<T>(ILogger<T> logger, M365AuthHelper auth) : ServiceBase<T>(logger)
{
    protected internal readonly M365AuthHelper Auth = auth;

    public async Task<IPnPContext> CreatePnPContextAsSystem() => await Auth.CreatePnPContextAsSystem();
    public async Task<IPnPContext> CreatePnPContextAsSystem(string siteRelativeUrl) => await Auth.CreatePnPContextAsSystem(siteRelativeUrl);
    public async Task<IPnPContext> CreatePnPContextAsSystem(IPnPContext ctx) => await Auth.CreatePnPContextAsSystem(ctx.Uri.AbsolutePath);
    public async Task<IPnPContext> CreatePnPContextAsUser(string accessToken) => await Auth.CreatePnPContextAsUser(accessToken);
    public async Task<IPnPContext> CreatePnPContextAsUser(string siteRelativeUrl, string accessToken) => await Auth.CreatePnPContextAsUser(siteRelativeUrl, accessToken);

    public async Task<IPnPContext> CloneToApprodotAsSystem(IPnPContext ctx) => await Auth.CloneToApprodotAsSystem(ctx);
    public async Task<IPnPContext> CloneTo(IPnPContext ctx, string siteRelativeUrl) => await Auth.CloneTo(ctx, siteRelativeUrl);
    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx, string siteRelativeUrl) => await Auth.CloneToAsSystem(ctx, siteRelativeUrl);
    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx) => await Auth.CloneToAsSystem(ctx);

    public async Task RunAsSystem(IPnPContext ctx, Func<IPnPContext, Task> action) => await Auth.RunAsSystem(ctx, ctx.Uri.AbsolutePath, action);
    public async Task RunAsSystem(IPnPContext ctx, string siteRelativeUrl, Func<IPnPContext, Task> action) => await Auth.RunAsSystem(ctx, siteRelativeUrl, action);
    public async Task RunAsSystemInApprodot(IPnPContext ctx, Func<IPnPContext, Task> action) => await Auth.RunAsSystemInApprodot(ctx, action);
    public async Task RunAsSystemInApprodot(Func<IPnPContext, Task> action) => await Auth.RunAsSystemInApprodot(action);

}
