using System;
using System.Threading.Tasks;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Xunit;
using Xunit.Abstractions;
using System.Diagnostics;
using Contoso.Portal.Common;
using Microsoft.Extensions.DependencyInjection;

namespace WebApi.Tests.Events;

public class EventMinutesTemplateServiceTest : BasePnPAppTest
{
    private string BODY_ID = "ma-gr-cyl";
    //private string SHARED_EVENT_ID = "2736cf55a16145b2b8371ccd98d2d941";
    private string SHARED_EVENT_ID = "c7400310fc644baa8abb5b7dea9df873";
    private readonly EventMinutesTemplateService _eventMinutesTemplateService;


    public EventMinutesTemplateServiceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        var bodyRoleService = new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth);
        var taxonomyService = new TaxonomyTranslationService(this.MemoryCache, Log<TaxonomyTranslationService>(), Auth, "es-ES", "contoso-dev-admin.sharepoint.com", "102a5d61-9d78-42e8-ab11-681331d4d37f");
        _eventMinutesTemplateService = Serviceprovider().GetRequiredService<EventMinutesTemplateService>();
    }



    [Fact]
    public async Task DownloadMinutesTemplateTest()
    {
        try
        {
            Debug.WriteLine("PRUEBA::: Inicio prueba generar Minutes.");
            using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);
            var response = await _eventMinutesTemplateService.DownloadMinutesTemplate(ctx, BODY_ID, SHARED_EVENT_ID);

            // $ ls /home/vscode/PTCAAPP/Azure/WebApi.Tests/bin/Debug/net6.0/prueba*
            // $ rm /home/vscode/PTCAAPP/Azure/WebApi.Tests/bin/Debug/net6.0/pruebaMinutes.docx
            string fileAct = @"./pruebaMinutes.docx";
            if (System.IO.File.Exists(fileAct)) try { System.IO.File.Delete(fileAct); } catch (Exception) { }
            string whereSave = System.IO.Directory.GetCurrentDirectory();
            Debug.WriteLine("PRUEBA::: localización del fichero Minutes generado= " + whereSave);
            System.IO.File.WriteAllBytes(fileAct, response);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }


}