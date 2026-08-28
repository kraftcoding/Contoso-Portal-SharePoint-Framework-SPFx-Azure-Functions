using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.OpenApi.Extensions;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Notifications;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Domains.DemoEntities;
using Contoso.Portal.Runtime;
using Contoso.Portal.Runtime.Middlewares;
using Contoso.Portal.Integrations;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services;
using PnP.Core.Services.Builder.Configuration;
using Contoso.Portal.Common;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.ApplicationInsights.WorkerService;
using Contoso.Portal.Domains.PowerBiReports;
using Contoso.Portal.Data.DAL.Helpers;


var host = new HostBuilder()
    .ConfigureFunctionsWebApplication((worker) =>
    {
        worker
        .UseNewtonsoftJson()
        .UseMiddleware<ExceptionHandlingMiddleware>();
    })
    .ConfigureOpenApi()
    .ConfigureAppConfiguration((context, configBuilder) =>
    {
        configBuilder
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddJsonFile($"appsettings.{context.HostingEnvironment.EnvironmentName}.json", optional: true, reloadOnChange: true)
            .AddEnvironmentVariables();
    })
    .ConfigureServices((context, services) =>
    {
        var settings = new Settings();
        var serviceProvider = services.BuildServiceProvider();
        var configuration = serviceProvider.GetService<IConfiguration>();
        configuration.Bind(settings);

        services.AddApplicationInsightsTelemetryWorkerService(new ApplicationInsightsServiceOptions()
        {
            EnableAdaptiveSampling = false,
        });
        services.ConfigureFunctionsApplicationInsights();
        services.Configure<LoggerFilterOptions>(options =>
        {
            var defaultFilterRules = options?.Rules?.Where(rule =>
                rule.ProviderName == "Microsoft.Extensions.Logging.ApplicationInsights.ApplicationInsightsLoggerProvider")
                .ToList() ?? [];
            foreach (var rule in defaultFilterRules)
                options?.Rules?.Remove(rule);
        });
        services.AddHttpClient();
        services.AddMemoryCache();
        services.AddSingleton(settings);

        // Mock integration services — replace with real providers for production
        services.AddSingleton<INotificationSmsService, MockNotificationSmsService>();
        services.AddSingleton<IDocumentArchiveService, MockDocumentArchiveService>();

        services.AddSingleton<DocumentToPdfService>();
        services.AddSingleton<BodyRoleService>();
        services.AddSingleton<BodiesService>();

        services.AddSingleton<EventMinutesService>();
        services.AddSingleton<EventMinutesTemplateService>();
        services.AddSingleton<EventAgendaItemsService>();
        services.AddSingleton<EventAttendanceService>();
        services.AddSingleton<EventAttendanceTemplateService>();
        services.AddSingleton<EventCaptureService>();
        services.AddSingleton<EventAgreementService>();
        services.AddSingleton<ProfileService>();
        services.AddSingleton<TasksService>();
        services.AddSingleton<TasksAttendanceService>();
        services.AddSingleton<TasksDelegateService>();
        services.AddSingleton<TasksApprovalMinutesService>();
        services.AddSingleton<TasksModificationMinutesService>();
        services.AddSingleton<TasksCertificationService>();
        services.AddSingleton<NotificationsBusinessService>();
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
            return new ConfigDepartmentsService(
                sp.GetRequiredService<IMemoryCache>(),
                sp.GetRequiredService<ILogger<ConfigDepartmentsService>>(),
                sp.GetRequiredService<M365AuthHelper>(),
                settings.DefaultLocale ?? "en-US",
                settings.TaxonomyRootSiteId ?? throw new Exception($"{nameof(Settings.TaxonomyRootSiteId)} not configured."),
                settings.TaxonomyAppTermGroup ?? throw new Exception($"{nameof(Settings.TaxonomyAppTermGroup)} not configured."),
                settings.TaxonomyDepartmentsSetId ?? throw new Exception($"{nameof(Settings.TaxonomyDepartmentsSetId)} not configured.")
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
        });

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
        });

        services.AddSingleton((sp) =>
        {
            return new DemoEntitiesService(
                sp.GetRequiredService<ILogger<DemoEntitiesService>>(),
                sp.GetRequiredService<M365AuthHelper>(),
                sp.GetRequiredService<DocumentToPdfService>(),
                sp.GetRequiredService<IDocumentArchiveService>(),
                sp.GetRequiredService<DataStorageHelper>()
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
                settings.DefaultLocale ?? "en-US",
                settings.GraphMailboxUserId ?? throw new Exception($"{nameof(Settings.GraphMailboxUserId)} not configured."));
        });
        services.AddSingleton((sp) =>
        {
            return new TasksDelegateService(
                sp.GetRequiredService<BodyRoleService>(),
                sp.GetRequiredService<ILogger<TasksDelegateService>>(),
                sp.GetRequiredService<M365AuthHelper>(),
                sp.GetRequiredService<IServiceProvider>(),
                sp.GetRequiredService<NotificationsService>(),
                sp.GetRequiredService<TasksAttendanceService>(),
                settings.DefaultLocale ?? "en-US"
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
                settings.DefaultLocale ?? "en-US",
                settings.TenantId ?? throw new Exception($"{nameof(Settings.TenantId)} not configured."),
                settings.Publishing_ReplicateSensitivityEnabled,
                settings.Publishing_PdfConversionEnabled);
        });
        services.AddSingleton((sp) =>
        {
            return new TaxonomyTranslationService(
                sp.GetRequiredService<IMemoryCache>(),
                sp.GetRequiredService<ILogger<TaxonomyTranslationService>>(),
                sp.GetRequiredService<M365AuthHelper>(),
                settings.DefaultLocale ?? "en-US",
                settings.TaxonomyRootSiteId ?? throw new Exception($"{nameof(Settings.TaxonomyRootSiteId)} not configured."),
                settings.TaxonomyAppTermGroup ?? throw new Exception($"{nameof(Settings.TaxonomyAppTermGroup)} not configured."));
        });
        services.AddSingleton((sp) =>
        {
            return new EventExchangeSyncService(
                sp.GetRequiredService<NotificationsService>(),
                sp.GetRequiredService<EventsService>(),
                sp.GetRequiredService<EventAttendanceService>(),
                sp.GetRequiredService<BodyRoleService>(),
                settings.GraphMailboxUserId ?? throw new Exception($"{nameof(Settings.GraphMailboxUserId)} not configured."),
                settings.DefaultLocale ?? "en-US",
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
            options.Sites.Add(DALConstants.PnPContexts.RootSiteAsSystem, new PnPCoreSiteOptions { SiteUrl = settings.RootSiteUrl });
            options.HttpRequests.UserAgent = "NONISV|CNT|Contoso/1.0";
            options.HttpRequests.Timeout = -1;
            options.DisableTelemetry = true;
        });
        services.AddPnPCoreAuthentication(options =>
        {
            options.Credentials.Configurations.Add("CertAuth", new PnPCoreAuthenticationCredentialConfigurationOptions
            {
                ClientId = settings.ClientId,
                TenantId = settings.TenantId,
                X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
                {
                    Certificate = M365AuthHelper.LoadCertificate(settings.CertificateThumbprint
                        ?? throw new Exception($"{nameof(Settings.CertificateThumbprint)} not configured."))
                }
            });
            options.Sites.Add(DALConstants.PnPContexts.RootSiteAsSystem, new PnPCoreAuthenticationSiteOptions { AuthenticationProviderName = "CertAuth" });
            options.Credentials.DefaultConfiguration = "CertAuth";
        });
    })
    .ConfigureLogging((hostingContext, logging) =>
    {
        logging.AddApplicationInsights(console =>
        {
            console.IncludeScopes = true;
        });
        logging.AddConfiguration(hostingContext.Configuration.GetSection("Logging"));
    })
    .Build();

host.Run();
