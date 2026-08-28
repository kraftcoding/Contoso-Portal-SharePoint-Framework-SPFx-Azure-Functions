using Contoso.Portal.Data.DAL;
using PnP.Core.Auth;
using PnP.Core.Services;
using System.Security.Cryptography.X509Certificates;
using Contoso.Portal.Data.Extensions;
using System.IdentityModel.Tokens.Jwt;
using Azure.Identity;
using Microsoft.Graph;

namespace Contoso.Portal.Common;

public class M365AuthHelper(string clientId, string tenantId, string certificateThumbprint, string approdotSiteUrl, IPnPContextFactory pnpContextFactory)
{
    private readonly string _clientId = clientId;
    private readonly string _tenantId = tenantId;
    private readonly string _certificateThumbprint = certificateThumbprint;

    public string ClientId => _clientId;
    public string TenantId => _tenantId;
    public string CertificateThumbprint => _certificateThumbprint;

    private readonly string _approdotServerRelativeUrl = new Uri(approdotSiteUrl ?? throw new ArgumentNullException(nameof(approdotSiteUrl))).AbsolutePath;
    private readonly IPnPContextFactory _pnpContextFactory = pnpContextFactory;
    private readonly string _sprodot = new Uri(approdotSiteUrl ?? throw new ArgumentNullException(nameof(approdotSiteUrl))).GetLeftPart(UriPartial.Authority);

    public IAuthenticationprovider GetOboprovider(string accessToken)
    {
        var certificate = LoadCertificate(_certificateThumbprint);
        return new OnBehalfOfAuthenticationprovider(_clientId, _tenantId, certificate, () => accessToken);
    }

    public IAuthenticationprovider GetApplicationprovider()
    {
        var certificate = LoadCertificate(_certificateThumbprint);
        return new X509CertificateAuthenticationprovider(_clientId, _tenantId, certificate);
    }

    public GraphServiceClient GetApplicationGraphClient()
    {
        var certificate = LoadCertificate(_certificateThumbprint);

        var credential = new ClientCertificateCredential(
            _tenantId,
            _clientId,
            certificate,
            new ClientCertificateCredentialOptions
            {
                SendCertificateChain = true
            }
        );

        return new GraphServiceClient(
            credential,
            new[] { "https://graph.microsoft.com/.default" }
        );
    }

    public async Task<IPnPContext> CreatePnPContextAppOnly(string siteRelativeUrl)
    {
        var ctx = await _pnpContextFactory.CreateAsync(
            new Uri(_sprodot.UriCombine(siteRelativeUrl)),
            GetApplicationprovider()
        );

        ctx.properties["CurrentUserUPN"] = DALConstants.SPAppLoginName;
        ctx.properties["IsElevated"] = true;

        return ctx;
    }

    public async Task<IPnPContext> CreatePnPContextAsSystem(string siteRelativeUrl)
    {
        var ctx = await _pnpContextFactory.CreateAsync(
            new Uri(_sprodot.UriCombine(siteRelativeUrl)),
            GetApplicationprovider()
        );

        ctx.properties["CurrentUserUPN"] = DALConstants.SPAppLoginName;
        ctx.properties["IsElevated"] = true;

        return ctx;
    }

    public async Task<IPnPContext> CreatePnPContextAsUser(string siteRelativeUrl, string accessToken)
    {
        var ctx = await _pnpContextFactory.CreateAsync(
            new Uri(_sprodot.UriCombine(siteRelativeUrl)),
            GetOboprovider(accessToken)
        );

        var claims = GetClaimsFromToken(accessToken);
        if (claims != null)
        {
            ctx.properties["CurrentUserUPN"] = claims.TryGetValue("upn", out var upn) ? upn : null;
            ctx.properties["AadObjectId"] = claims.TryGetValue("oid", out var oid) ? oid : null;
            ctx.properties["IsElevated"] = false;
        }

        return ctx;
    }

    public async Task<IPnPContext> CloneTo(IPnPContext ctx, string siteRelativeUrl)
        => await ctx.CloneAsync(new Uri(_sprodot.UriCombine(siteRelativeUrl)));

    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx, string siteRelativeUrl)
    {
        var sameSite = ctx.Uri.AbsolutePath.Trim('/').Equals(siteRelativeUrl.Trim('/'), StringComparison.OrdinalIgnoreCase);
        var elevated = await ctx.IsElevatedOrSystem();

        if (elevated)
            return sameSite ? await ctx.CloneAsync() : await ctx.CloneAsync(new Uri(_sprodot.UriCombine(siteRelativeUrl)));

        var sys = await CreatePnPContextAsSystem(siteRelativeUrl);
        sys.properties["CurrentUserUPN"] = await ctx.GetCurrentUserUpn(false);
        sys.properties["AadObjectId"] = await ctx.GetCurrentUserId(false);
        return sys;
    }

    public async Task RunAsSystem(string siteRelativeUrl, Func<IPnPContext, Task> action)
    {
        using var ctx = await CreatePnPContextAsSystem(siteRelativeUrl);
        await action(ctx);
    }

    public async Task RunAsSystem(IPnPContext ctx, string siteRelativeUrl, Func<IPnPContext, Task> action)
    {
        var sameSite = ctx.Uri.AbsolutePath.Trim('/').Equals(siteRelativeUrl.Trim('/'), StringComparison.OrdinalIgnoreCase);
        var elevated = await ctx.IsElevatedOrSystem();

        if (elevated && sameSite)
        {
            await action(ctx);
            return;
        }

        using var sys = elevated
            ? await ctx.CloneAsync(new Uri(_sprodot.UriCombine(siteRelativeUrl)))
            : await CloneToAsSystem(ctx, siteRelativeUrl);

        await action(sys);
    }

    public async Task<IPnPContext> CreatePnPContextAsSystem()
        => await CreatePnPContextAsSystem(_approdotServerRelativeUrl);

    public async Task<IPnPContext> CreatePnPContextAsUser(string accessToken)
        => await CreatePnPContextAsUser(_approdotServerRelativeUrl, accessToken);

    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx)
        => await CloneToAsSystem(ctx, ctx.Uri.AbsolutePath);

    public async Task RunAsSystem(IPnPContext ctx, Func<IPnPContext, Task> action)
        => await RunAsSystem(ctx, ctx.Uri.AbsolutePath, action);

    public async Task RunAsSystemInApprodot(IPnPContext ctx, Func<IPnPContext, Task> action)
        => await RunAsSystem(ctx, _approdotServerRelativeUrl, action);

    public async Task RunAsSystemInApprodot(Func<IPnPContext, Task> action)
        => await RunAsSystem(_approdotServerRelativeUrl, action);

    public async Task<IPnPContext> CloneToApprodotAsSystem(IPnPContext ctx)
        => await CloneToAsSystem(ctx, _approdotServerRelativeUrl);

    public static X509Certificate2 LoadCertificate(string certificateThumbprint)
    {
        var mode = Environment.GetEnvironmentVariable("CertificateLoadMode") ?? null;

        if (mode == "store" || mode == null)
        {
            var thumbprint = Environment.GetEnvironmentVariable(certificateThumbprint)
                ?? throw new InvalidOperationException("CertificateThumbPrint not configured.");

            return LoadCertificateFromCertStore(thumbprint);
        }
        else
        {
            var fileName = Environment.GetEnvironmentVariable("CertificateFileName") ?? "auth.vhorstyle.com.pfx";
            var password = Environment.GetEnvironmentVariable("CertificatePassword") ?? "rg-dev-contoso";
            return LoadCertificateFromFile(fileName, password);
        }
    }

    public static X509Certificate2 LoadCertificateFromCertStore(string certificateThumbprint)
    {
        string? certBase64 = Environment.GetEnvironmentVariable("CertificateFromKeyVault");

        if (!string.IsNullOrECNTy(certBase64))
        {
            return new X509Certificate2(
                Convert.FromBase64String(certBase64),
                "",
                X509KeyStorageFlags.Exportable |
                X509KeyStorageFlags.MachineKeySet |
                X509KeyStorageFlags.EphemeralKeySet
            );
        }

        var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        var certs = store.Certificates.Find(X509FindType.FindByThumbprint, certificateThumbprint, false);
        store.Close();

        return certs.First();
    }

    public static X509Certificate2 LoadCertificateFromFile(string fileName, string password)
    {
        var home = Environment.GetEnvironmentVariable("HOME")
            ?? Environment.GetEnvironmentVariable("USERprodFILE");

        var certPath = Path.Combine(home, "certificados", fileName);

        if (!File.Exists(certPath))
            throw new FileNotFoundException($"Certificate not found at: {certPath}");

        return new X509Certificate2(
            certPath,
            password,
            X509KeyStorageFlags.MachineKeySet |
            X509KeyStorageFlags.Exportable |
            X509KeyStorageFlags.EphemeralKeySet
        );
    }

    public static IDictionary<string, string>? GetClaimsFromToken(string? accessToken)
    {
        if (string.IsNullOrECNTy(accessToken))
            return null;

        var handler = new JwtSecurityTokenHandler();
        var token = handler.ReadToken(accessToken) as JwtSecurityToken;
        var result = new Dictionary<string, string>();

        foreach (var claim in token?.Claims ?? [])
            result.TryAdd(claim.Type, claim.Value);

        return result;
    }

    public static string? GetUpnFromToken(string? accessToken)
    {
        var claims = GetClaimsFromToken(accessToken);
        return claims?.TryGetValue("upn", out var upn) == true ? upn : null;
    }

    public static string? GetOidFromToken(string? accessToken)
    {
        var claims = GetClaimsFromToken(accessToken);
        return claims?.TryGetValue("oid", out var oid) == true ? oid : null;
    }
}
