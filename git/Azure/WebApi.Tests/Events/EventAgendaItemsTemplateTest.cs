using System;
using System.Threading.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Xunit;
using Xunit.Abstractions;
using System.Diagnostics;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Common;
using System.IO;

namespace WebApi.Tests.Events;

public class EventAgendaItemsTemplateTest : BasePnPAppTest
{
    private string BODY_ID = "ma-gr-cyl";
    private string SHARED_EVENT_ID = "47ca7eee13904872b128b1fa185feef4";
    private EventAgendaItemsService _eventAgendaItemsService;
    private readonly TaxonomyTranslationService _taxonomyService;

    private readonly EventsService _eventsService;


    public EventAgendaItemsTemplateTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        var roleSrv = new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth);
        var documentToPDFService = ServiceProvider().GetRequiredService<DocumentToPdfService>();
        _eventsService = ServiceProvider().GetRequiredService<EventsService>();
        _taxonomyService = new TaxonomyTranslationService(this.MemoryCache, Log<TaxonomyTranslationService>(), Auth, "es-ES", "contoso-dev-admin.sharepoint.com", "102a5d61-9d78-42e8-ab11-681331d4d37f");
        _eventAgendaItemsService = new EventAgendaItemsService(roleSrv,
                                                            new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth, "es-ES", "contoso-dev-admin.sharepoint.com", "102a5d61-9d78-42e8-ab11-681331d4d37f", ""),
                                                            // new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth),
                                                            _eventsService,
                                                            _taxonomyService,
                                                            Log<EventAgendaItemsService>(), Auth, documentToPDFService
                                                           ); ;
    }


    [Fact]
    public async Task GetAgendaItemsTemplateTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var result = await _eventAgendaItemsService.DownloadAgendaItemsTemplate(ctx, BODY_ID, SHARED_EVENT_ID);
            string ficheroGuarda = @"./pruebaOrden.docx";

            if (System.IO.File.Exists(ficheroGuarda)) try { System.IO.File.Delete(ficheroGuarda); } catch (Exception) { }

            string whereSave = System.IO.Directory.GetCurrentDirectory();
            Debug.WriteLine("PRUEBA::: localización del fichero generado= " + whereSave);
            File.WriteAllBytes(ficheroGuarda, result);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }
}