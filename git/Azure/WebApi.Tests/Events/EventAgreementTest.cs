using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Model.Events;
using Xunit;
using Xunit.Abstractions;

namespace WebApi.Tests.Events;

public class EventAgreementTest : BasePnPAppTest
{
    private string BODY_ID = "demoev-cs-voto";
    //private string SHARED_EVENT_ID = "bf6621c9f9814092bdb4138ad3550d10"; // [1090 a 1091]
    private string SHARED_EVENT_ID = "a5dc0d02949842f9ae890bdfe928c88e"; // 974
    private string AGREEMENT_ID = "477bc7fc0d534a18a9af2557e422906d";

    private object _eventAgreements;
    private readonly EventAgreementService _eventAgreementService;

    private readonly TaxonomyTranslationService _taxonomyService;

    private readonly EventsService _eventsService;
    private readonly EventVotationService _eventVotationService;


    public EventAgreementTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        var roleSrv = new BodyRoleService(this.MemoryCache, Log<BodyRoleService>(), Auth);
        _eventsService = ServiceProvider().GetRequiredService<EventsService>();
        _eventVotationService = ServiceProvider().GetRequiredService<EventVotationService>();
        _taxonomyService = new TaxonomyTranslationService(this.MemoryCache, Log<TaxonomyTranslationService>(), Auth, "es-ES", "contoso-dev-admin.sharepoint.com", "102a5d61-9d78-42e8-ab11-681331d4d37f");
        _eventAgreementService = new EventAgreementService(roleSrv,
                                                            new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth, "es-ES", "contoso-dev-admin.sharepoint.com", "102a5d61-9d78-42e8-ab11-681331d4d37f", ""),
                                                            // new ConfigDepartmentsService(this.MemoryCache, Log<ConfigDepartmentsService>(), Auth),
                                                            _eventsService,
                                                            _taxonomyService,
                                                            _eventVotationService,
                                                            Log<EventAgreementService>(), Auth);
        _eventAgreements = new EventAgreement();
    }

    [Fact]
    public async Task GetAgreements()
    {
        using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);

        var result = ElapsedUtils.Time(async () =>
        {
            try
            {
                var agreements = await _eventAgreementService.GetAgreementsByEvent(ctx, BODY_ID, SHARED_EVENT_ID);
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }, 5);

        _eventAgreements = new EventAgreement();
    }




    [Fact]
    public async Task DownloadAgreementsTemplateTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID); //using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var response = await _eventAgreementService.DownloadAgreementsTemplate(ctx, BODY_ID, SHARED_EVENT_ID, AGREEMENT_ID);

            string ficheroGuarda = @"./pruebaAcuerdo.docx";
            // $ ls /home/vscode/PTCAAPP/Azure/WebApi.Tests/bin/Debug/net6.0/prueba.*
            // $ rm /home/vscode/PTCAAPP/Azure/WebApi.Tests/bin/Debug/net6.0/prueba.docx
            if (System.IO.File.Exists(ficheroGuarda)) try { System.IO.File.Delete(ficheroGuarda); } catch (Exception) { }

            string whereSave = System.IO.Directory.GetCurrentDirectory();
            Debug.WriteLine("PRUEBA::: localización del fichero generado= " + whereSave);
            File.WriteAllBytes(ficheroGuarda, response);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateAgreements()
    {
        try
        {
            EventAgreement[] eventAgreementsArray = (EventAgreement[])_eventAgreements;
            List<EventAgreement> result = new List<EventAgreement>
            {
                new() { Id = "0", },
            };
            using var ctx = await CreatePnPContextAsUser("sites/" + BODY_ID);
            var agreements = await _eventAgreementService.AddOrUpdateAgreementsForEvent(ctx, BODY_ID, SHARED_EVENT_ID, eventAgreementsArray);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }

    [Fact]
    public async Task GetAgreementTemplate()
    {
        try
        {

            using var ctx = await CreatePnPContextAsSystem("sites/" + BODY_ID);
            var agreements = await _eventAgreementService.DownloadAgreementsTemplate(ctx, BODY_ID, SHARED_EVENT_ID, AGREEMENT_ID);

        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }

    }
}