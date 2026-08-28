using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Domains.Tasks;

using Contoso.Portal.Runtime;
using PnP.Core.Auth.Services;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

using System.Diagnostics;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Data.DAL.Helpers;


public class InitializedHostFixture : IDisposable
{
    public IHost TestHost { get; }

    public InitializedHostFixture()
    {
        TestHost = CreateHostBuilder().Build();
        Task.Run(() => TestHost.RunAsync());
    }

    public static IHostBuilder CreateHostBuilder() =>
        Host.CreateDefaultBuilder()
            .ConfigureServices((context, services) =>
            {
                var settings = new Settings();
                context.Configuration.Bind(settings);
                services.AddSingleton(options => { return settings; });
                services.AddHttpClient();
                services.AddMemoryCache();
                services.AddSingleton<DocumentToPdfService>();
                services.AddSingleton<BodyRoleService>();
                services.AddSingleton<BodiesService>();
                // services.AddSingleton<ConfigDepartmentsService>();
                services.AddSingleton((sp) =>
                {
                    return new ConfigDepartmentsService(
                        sp.GetRequiredService<IMemoryCache>(),
                        sp.GetRequiredService<ILogger<ConfigDepartmentsService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        settings.DefaultLocale ?? DALConstants."en-US",
                        settings.TaxonomyRootSiteId ?? throw new Exception($"{nameof(Settings.TaxonomyRootSiteId)} not configured."),
                        settings.TaxonomyAppTermGroup ?? throw new Exception($"{nameof(Settings.TaxonomyAppTermGroup)} not configured."),
                        settings.TaxonomyDepartmentsSetId ?? throw new Exception($"{nameof(Settings.TaxonomyDepartmentsSetId)} not configured.")
                    );
                });

                services.AddSingleton<EventMinutesService>();
                services.AddSingleton<EventMinutesTemplateService>();
                services.AddSingleton<EventAttendanceService>();
                services.AddSingleton<EventAttendanceTemplateService>();
                services.AddSingleton<EventCaptureService>();
                services.AddSingleton<EventAgreementService>();
                services.AddSingleton<ProfileService>();
                services.AddSingleton<TasksService>();
                services.AddSingleton<TasksAttendanceService>();
                services.AddSingleton<TasksApprovalMinutesService>();
                services.AddSingleton<TasksModificationMinutesService>();
                services.AddSingleton<TasksCertificationService>();
                services.AddSingleton<NotificationsBusinessService>();
                services.AddSingleton<EventAgendaItemsService>();
                services.AddSingleton((sp) =>
                {
                    return new EventsService(
                        sp.GetRequiredService<BodyRoleService>(),
                        sp.GetRequiredService<ILogger<EventsService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        sp.GetRequiredService<IServiceProvider>(),
                        sp.GetRequiredService<ConfigDepartmentsService>(),
                        settings.VotingEnabled ?? false
                    );
                });

                services.AddSingleton((sp) =>
                {
                    return new EventVotationService(
                        sp.GetRequiredService<BodyRoleService>(),
                        sp.GetRequiredService<ProfileService>(),
                        sp.GetRequiredService<EventAttendanceService>(),
                        sp.GetRequiredService<ILogger<EventVotationService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        settings.TaxonomyAppTermGroup ?? throw new Exception($"{nameof(Settings.TaxonomyAppTermGroup)} not configured."),
                        settings.TaxonomyRootSiteId ?? throw new Exception($"{nameof(Settings.TaxonomyRootSiteId)} not configured."),
                        settings.VotingEnabled ?? false
                    );
                }
                );

                services.AddSingleton((sp) =>
       {
           return new ManagementService(
               sp.GetRequiredService<IMemoryCache>(),
               sp.GetRequiredService<ILogger<ManagementService>>(),
               sp.GetRequiredService<M365AuthHelper>(),
               sp.GetRequiredService<ConfigDepartmentsService>(),
               sp.GetRequiredService<BodyRoleService>(),
               sp.GetRequiredService<DocumentToPdfService>(),
               sp.GetRequiredService<IDocumentArchiveService>()
           );
       }
       );

                services.AddSingleton((sp) =>
                {
                    return new TasksDelegateService(
                        sp.GetRequiredService<BodyRoleService>(),
                        sp.GetRequiredService<ILogger<TasksDelegateService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        sp.GetRequiredService<IServiceProvider>(),
                        sp.GetRequiredService<NotificationsService>(),
                        sp.GetRequiredService<TasksAttendanceService>(),
                        settings.DefaultLocale ?? DALConstants."en-US"
                    );
                });
                services.AddSingleton((sp) =>
                {
                    return new NotificationsService(
                        sp.GetRequiredService<BodyRoleService>(),
                        sp.GetRequiredService<INotificationSmsService>(),
                        sp.GetRequiredService<ProfileService>(),
                        sp.GetRequiredService<ConfigDepartmentsService>(),
                        sp.GetRequiredService<ILogger<NotificationsService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        sp.GetRequiredService<IMemoryCache>(),
                        sp.GetRequiredService<IServiceProvider>(),
                        settings.DefaultLocale ?? DALConstants."en-US",
                        settings.GraphMailboxUserId ?? throw new Exception($"{nameof(Settings.GraphMailboxUserId)} not configured."));
                });

                services.AddSingleton((sp) =>
                {
                    return new PowerBiReportsService(
                        sp.GetRequiredService<ILogger<PowerBiReportsService>>(),
                        sp.GetRequiredService<M365AuthHelper>(),
                        sp.GetRequiredService<DataStorageHelper>(),
                        sp.GetRequiredService<BodyRoleService>(),
                        sp.GetRequiredService<ConfigDepartmentsService>(),
                        sp.GetRequiredService<EventMinutesService>(),
                        sp.GetRequiredService<EventAttendanceService>()
                    );
                });

                services.AddSingleton((sp) =>
                {
                    return new M365AuthHelper(
                        settings.ClientId ?? throw new Exception($"{nameof(Settings.ClientId)} not configured."),
                        settings.TenantId ?? throw new Exception($"{nameof(Settings.TenantId)} not configured."),
                        settings.CertificateThumbprint ?? throw new Exception($"{nameof(Settings.CertificateThumbprint)} not configured."),
                        settings.RootSiteUrl ?? throw new Exception($"{nameof(Settings.RootSiteUrl)} not configured."),
                        sp.GetRequiredService<IPnPContextFactory>());
                });
                services.AddSingleton((sp) =>
                {
                    return new DataStorageHelper(
                        settings.ClientId ?? throw new Exception($"{nameof(Settings.ClientId)} not configured."),
                        settings.TenantId ?? throw new Exception($"{nameof(Settings.TenantId)} not configured."),
                        M365AuthHelper.LoadCertificate(settings.CertificateThumbprint ?? throw new Exception($"{nameof(Settings.CertificateThumbprint)} not configured.")),
                        settings.StorageAccountName ?? throw new Exception($"{nameof(Settings.StorageAccountName)} not configured"));
                });
                services.AddSingleton((sp) =>
                {
                    return new EventPublishingService(
                            sp.GetRequiredService<BodyRoleService>(),
                            sp.GetRequiredService<NotificationsService>(),
                            sp.GetRequiredService<TasksService>(),
                            sp.GetRequiredService<IDocumentArchiveService>(),
                            sp.GetRequiredService<ConfigDepartmentsService>(),
                            sp.GetRequiredService<EventExchangeSyncService>(),
                            sp.GetRequiredService<ILogger<EventPublishingService>>(),
                            sp.GetRequiredService<M365AuthHelper>(),
                            settings.GraphMailboxUserId ?? throw new Exception($"{nameof(Settings.GraphMailboxUserId)} not configured."),
                            settings.DefaultLocale ?? DALConstants."en-US",
                            settings.TenantId ?? throw new Exception($"{nameof(Settings.TenantId)} not configured."),
                            settings.Publishing_ReplicateSensitivityEnabled,
                            settings.Publishing_PdfConversionEnabled);
                });
                services.AddSingleton((sp) =>
                {
                    return new TaxonomyTranslationService(
                    sp.GetRequiredService<IMemoryCache>()
                    , sp.GetRequiredService<ILogger<TaxonomyTranslationService>>()
                    , sp.GetRequiredService<M365AuthHelper>()
                    , settings.DefaultLocale ?? DALConstants."en-US"
                    , settings.TaxonomyRootSiteId ?? throw new Exception($"{nameof(Settings.TaxonomyRootSiteId)} not configured.")
                    , settings.TaxonomyAppTermGroup ?? throw new Exception($"{nameof(Settings.TaxonomyAppTermGroup)} not configured."));
                });
                services.AddSingleton((sp) =>
                {
                    return new INotificationSmsService(
                          sp.GetRequiredService<ILogger<INotificationSmsService>>()
                        , settings.SmsServiceUser ?? throw new Exception($"{nameof(Settings.SmsServiceUser)} not configured.")
                        , settings.SmsServicePassword ?? throw new Exception($"{nameof(Settings.SmsServicePassword)} not configured.")
                        , settings.SmsServiceId ?? throw new Exception($"{nameof(Settings.SmsServiceId)} not configured.")
                        , settings.SmsServiceEnabled ?? throw new Exception($"{nameof(Settings.SmsServiceEnabled)} not configured.")
                    );
                });
                services.AddSingleton((sp) =>
                {
                    return new IDocumentArchiveService(
                        sp.GetRequiredService<ILogger<IDocumentArchiveService>>()
                        , settings.ArchiveServiceUser ?? throw new Exception($"{nameof(Settings.ArchiveServiceUser)} not configured.")
                        , settings.ArchiveServicePassword ?? throw new Exception($"{nameof(Settings.ArchiveServicePassword)} not configured.")
                    );
                });
                services.AddSingleton((sp) =>
                {
                    return new EventExchangeSyncService(
                        sp.GetRequiredService<NotificationsService>(),
                        sp.GetRequiredService<EventsService>(),
                        sp.GetRequiredService<EventAttendanceService>(),
                        sp.GetRequiredService<BodyRoleService>(),
                        settings.GraphMailboxUserId ?? throw new Exception($"{nameof(Settings.GraphMailboxUserId)} not configured."),
                        settings.DefaultLocale ?? DALConstants."en-US",
                        sp.GetRequiredService<ILogger<EventExchangeSyncService>>(),
                        sp.GetRequiredService<M365AuthHelper>()
                    );
                });
                services.AddSingleton((sp) =>
                {
                    return new EventFilesSensitivityLabelsService(
                        sp.GetRequiredService<BodyRoleService>(),
                        settings.TenantId!,
                        sp.GetRequiredService<ILogger<EventFilesSensitivityLabelsService>>(),
                        sp.GetRequiredService<M365AuthHelper>()
                    );
                });

                // Add and configure PnP Core SDK
                services.AddPnPCore(options =>
                {
                    // Add the base site url
                    options.Sites.Add(DALConstants.PnPContexts.RootSiteAsSystem, new PnPCoreSiteOptions { SiteUrl = settings.RootSiteUrl });
                    options.DisableTelemetry = true;
                });
                services.AddPnPCoreAuthentication(options =>
                {
                    // Configure certificate based auth and set it as default
                    options.Credentials.Configurations.Add("CertAuth", new PnPCoreAuthenticationCredentialConfigurationOptions
                    {
                        ClientId = settings.ClientId,
                        TenantId = settings.TenantId,
                        X509Certificate = new PnPCoreAuthenticationX509CertificateOptions { Certificate = M365AuthHelper.LoadCertificate(settings.CertificateThumbprint) }
                    });
                    options.Sites.Add(DALConstants.PnPContexts.RootSiteAsSystem, new PnPCoreAuthenticationSiteOptions { AuthenticationProviderName = "CertAuth" });
                    options.Credentials.DefaultConfiguration = "CertAuth";

                    // Configure an Authentication provider relying on Windows Credential Manager
                    options.Credentials.Configurations.Add("InteractiveAuth",
                        new PnPCoreAuthenticationCredentialConfigurationOptions
                        {
                            ClientId = settings.ClientId,
                            TenantId = settings.TenantId,
                            Interactive = new PnPCoreAuthenticationInteractiveOptions
                            {
                                RedirectUri = new Uri("http://localhost")
                            }
                        });
                });
            })
            .ConfigureLogging((logging) =>
            {
                logging.AddConsole();
            });

    public void Dispose()
    {
        this.TestHost.Dispose();
    }
}

[CollectionDefinition("Initialized host collection")]
public class PnPHostCollection : ICollectionFixture<InitializedHostFixture>
{
    // This class has no code, and is never created. Its purpose is simply
    // to be the place to apply [CollectionDefinition] and all the
    // ICollectionFixture<> interfaces.
}

[Collection("Initialized host collection")]
public class BasePnPAppTest
{
    InitializedHostFixture fixture;

    public ILogger<T> Log<T>() => this.fixture.TestHost.Services.GetRequiredService<ILogger<T>>();
    public ILoggerFactory LogFactory() => this.fixture.TestHost.Services.GetRequiredService<ILoggerFactory>();
    public IServiceProvider ServiceProvider() => this.fixture.TestHost.Services.GetRequiredService<IServiceProvider>();
    public IPnPContextFactory PnPContextFactory { get; }
    public M365AuthHelper Auth { get; }
    public Settings Settings { get; }
    public HttpClient HttpClient { get; }
    public IMemoryCache MemoryCache { get; }
    public readonly ITestOutputHelper Output;

    public BasePnPAppTest(InitializedHostFixture fixture, ITestOutputHelper output)
    {
        this.fixture = fixture;
        this.Output = output;
        this.PnPContextFactory = this.fixture.TestHost.Services.GetRequiredService<IPnPContextFactory>();
        this.Auth = this.fixture.TestHost.Services.GetRequiredService<M365AuthHelper>();
        this.Settings = this.fixture.TestHost.Services.GetRequiredService<Settings>();
        this.HttpClient = this.fixture.TestHost.Services.GetRequiredService<HttpClient>();
        this.MemoryCache = this.fixture.TestHost.Services.GetRequiredService<IMemoryCache>();
    }

    public IHost TestHost { get { return this.fixture.TestHost; } }

    public async Task<IPnPContext> CreatePnPContextAsUser(string siteRelativeUrl, string accessToken) => await Auth.CreatePnPContextAsUser(siteRelativeUrl, accessToken);
    public async Task<IPnPContext> CreatePnPContextAsSystem() => await Auth.CreatePnPContextAsSystem();
    public async Task<IPnPContext> CreatePnPContextAsSystem(string siteRelativeUrl) => await Auth.CreatePnPContextAsSystem(siteRelativeUrl);

    public async Task<IPnPContext> CreatePnPContextAsUser()
    {
        var interactiveAuth = this.TestHost.Services.GetRequiredService<IAuthenticationproviderFactory>().Create("InteractiveAuth");
        return await this.PnPContextFactory.CreateAsync(new Uri(this.Settings.RootSiteUrl ?? throw new ArgumentNullException(nameof(this.Settings.RootSiteUrl))), interactiveAuth);
    }

    public async Task<IPnPContext> CreatePnPContextAsUser(string siteRelativeUrl)
    {
        var interactiveAuth = this.TestHost.Services.GetRequiredService<IAuthenticationproviderFactory>().Create("InteractiveAuth");
        var sprodotUrl = new Uri(this.Settings.RootSiteUrl ?? throw new ArgumentNullException(nameof(this.Settings.RootSiteUrl))).GetLeftPart(UriPartial.Authority);
        var spSiteUrl = new Uri(sprodotUrl.UriCombine(siteRelativeUrl));

        return await this.PnPContextFactory.CreateAsync(spSiteUrl, interactiveAuth);
    }


    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx, string siteRelativeUrl)
    {
        var elevated = await CreatePnPContextAsSystem(siteRelativeUrl);
        elevated.properties.Add("CurrentUserUPN", await ctx.GetCurrentUserUpn(false));
        return elevated;
    }

    public async Task<IPnPContext> CloneToAsSystem(IPnPContext ctx) => await CloneToAsSystem(ctx, ctx.Uri.LocalPath);

    public async Task RunAsSystem(IPnPContext ctx, string siteRelativeUrl, Action<IPnPContext> action)
    {
        using var elevatedCtx = await CreatePnPContextAsSystem(siteRelativeUrl);
        elevatedCtx.properties.Add("CurrentUserUPN", await ctx.GetCurrentUserUpn(false));
        action(elevatedCtx);
    }

    public async Task RunAsSystem(IPnPContext ctx, string siteRelativeUrl, Func<IPnPContext, Task> action)
    {
        using var elevatedCtx = await CreatePnPContextAsSystem(siteRelativeUrl);
        elevatedCtx.properties.Add("CurrentUserUPN", await ctx.GetCurrentUserUpn(false));
        await action(elevatedCtx);
    }

    public async Task RunAsSystem(IPnPContext ctx, Action<IPnPContext> action) => await RunAsSystem(ctx, ctx.Uri.LocalPath, action);
    public async Task RunAsSystem(IPnPContext ctx, Func<IPnPContext, Task> action) => await RunAsSystem(ctx, ctx.Uri.LocalPath, action);

}

public static class ElapsedUtils
{
    public static TimeSpan Time(Action action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public static TimeSpan Time(Func<Task> action)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        action().Wait();
        stopwatch.Stop();
        return stopwatch.Elapsed;
    }

    public static double Time(Func<Task> action, int executions)
    {
        var loops = executions;

        Stopwatch stopwatch = Stopwatch.StartNew();
        while (--executions > 0)
        {
            action().Wait();
        }
        stopwatch.Stop();

        return stopwatch.Elapsed.TotalMilliseconds / loops;
    }
}