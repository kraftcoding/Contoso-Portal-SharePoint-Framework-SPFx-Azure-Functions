using Microsoft.Extensions.DependencyInjection;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.profile;
using System;
using System.Threading.Tasks;
using System.Linq;
using Xunit;
using Xunit.Abstractions;
using Newtonsoft.Json;
using Contoso.Portal.Data.DAL.Helpers;
using System.Collections.Generic;
using Contoso.Portal.Data.DAL.Management;
using System.IO;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Microsoft.Azure.WebJobs.Extensions.OpenApi.Core.Extensions;
using Contoso.Portal.Model.Management;
using Contoso.Portal.Data.DAL.ConfigDepartments;

namespace WebApi.Tests;

public class ManagementServiceTest : BasePnPAppTest
{
    private string bodyId = "Contoso";

    private string userUPN = "contoso-testadmin@contoso-test.onmicrosoft.com";
    private string memberassistantupn = "CNTextasistentemiembro@contoso-test.onmicrosoft.com";

    private string siteId = "contoso-test-admin.sharepoint.com";
    private string groupId = "102a5d61-9d78-42e8-ab11-681331d4d37f";
    private ManagementService _managementService;
    private ConfigDepartmentsService _ConfigDepartmentsService;

    public ManagementServiceTest(InitializedHostFixture fixture, ITestOutputHelper output) : base(fixture, output)
    {

        _managementService = ServiceProvider().GetRequiredService<ManagementService>();
        _ConfigDepartmentsService = ServiceProvider().GetRequiredService<ConfigDepartmentsService>();
    }

    // [Fact]
    // public async Task DownloadCertificateTemplateTest()
    // {
    //     try
    //     {
    //         using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");
    //         var arrBytes = await _managementService.DownloadCertificateTemplate(ctx, userUPN, "Contoso", "demoev-cs-voto");
    //         File.WriteAllBytes("Foo.pdf", arrBytes);
    //         Assert.True(true);
    //     }
    //     catch (Exception ex)
    //     {
    //         Assert.Fail(ex.Message);
    //     }
    // }

    [Fact]
    public async Task DownloadEXTERNALCertificateTemplateTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");
            var arrBytes = await _managementService.DownloadEXTERNALCertificateTemplate(ctx, userUPN, "Contoso", "AAA20150002");
            File.WriteAllBytes("Foo.pdf", arrBytes);
            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetConfigDepartments()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");
            //var siteId = ctx.Site.Id.ToString();

            var bodies = await _managementService.GetBodiesInformation(ctx, userUPN, bodyId);

            var jsonOb = JsonConvert.SerializeObject(bodies.First());




            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetEXTERNALConfigDepartmentsTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var siteId = ctx.Site.Id.ToString();

            var bodies = await _managementService.GetEXTERNALBodiesInformation(ctx, userUPN, bodyId);

            var jsonOb = JsonConvert.SerializeObject(bodies.First());


            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task AddBodyUsersRolesRequestTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            List<UserBodyRole> list = [new UserBodyRole() {
                UserUpn = "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com",
                BodyId = "demoev-co-cna",
                Role = "asistentemembers"
            },new UserBodyRole() {
                UserUpn = "kmartinr_emeal.nttdata.com@contoso-test.onmicrosoft.com",
                BodyId = "demomvp-cs-dep",
                Role = "gestorschedulers"
            },new UserBodyRole() {
                UserUpn = "dsancbar_emeal.nttdata.com@contoso-test.onmicrosoft.com",
                BodyId = "demomvp-cs-ae",
                Role = "asistentemembers"
            }];

            await _managementService.AddBodyUsersRolesRequest(ctx, userUPN, bodyId, list);
            Assert.True(true);
        }
        catch (System.Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task CheckPermissionsTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/demoev-cs-voto");
            var result = await _managementService.CheckPermissions(ctx, new UserBodyRole()
            {
                BodyId = "demoev-cs-voto",
                Role = "members",
                UserUpn = "CNTextmiembro1@contoso-test.onmicrosoft.com"
            });

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetAvailableLicensesTest()
    {
        try
        {
            bodyId = "Contoso";
            userUPN = "contoso-testadmin@contoso-test.onmicrosoft.com";
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var result = await _managementService.GetAvailableLicenses(ctx, userUPN, bodyId);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task ManageUserLicenseTest()
    {
        try
        {
            bodyId = "Contoso";
            userUPN = "contoso-testadmin@contoso-test.onmicrosoft.com";
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var userInfoAdd = System.Text.Json.JsonSerializer.Deserialize<UserInfo>("{\n    \"AssignedLicenses\": [\n        \"c42b9cae-ea4f-4ab7-9717-81576235ccac\"\n    ],\n    \"UserPrincipalName\": \"CNTextmiembro2@contoso-test.onmicrosoft.com\"\n}");
            var userInfoRemove = System.Text.Json.JsonSerializer.Deserialize<UserInfo>("{\n    \"AssignedLicenses\": [\n   ],\n    \"UserPrincipalName\": \"CNTextmiembro2@contoso-test.onmicrosoft.com\"\n}");
            var result = await _managementService.ManageUserLicense(ctx, userUPN, bodyId, userInfoAdd!);

            Assert.NotNull(result);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetUserInfo()
    {
        try
        {
            bodyId = "Contoso";
            userUPN = "contoso-testadmin@contoso-test.onmicrosoft.com";
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");

            var users = await _managementService.GetUsersInformation(ctx, userUPN, bodyId);

            // var hola = System.Text.Json.JsonSerializer.Deserialize<UserInfo>("{\n    \"AssignedLicenses\": [\n        \"f30db892-07e9-47e9-837c-80727f46fd3d\",\n        \"c42b9cae-ea4f-4ab7-9717-81576235ccac\"\n    ],\n    \"UserPrincipalName\": \"comunicacion@contoso-test.onmicrosoft.com\"\n}");

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task CreateNewUserTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");

            var users = await _managementService.CreateUser(ctx, userUPN, bodyId, new()
            {
                UserBodyRole = new()
                {
                    BodyId = "demomvp-cs-ae",
                    Role = "schedulers",
                },
                UserInfo = new()
                {
                    FirstName = "Test2",
                    LastName = "Testez Testado",
                    PrincipalMail = "test@email.com",
                    Email = "othermail@test.com",
                    JobTitle = "Analista",
                    Office = "Madrid",
                    BusinessPhone = "777777777",
                    CellPhone = "789987879",
                    CompanyName = "NTT Data",
                    Department = "Modern Application Services",
                    EmployeeType = "Informático",
                },
                UserInvitation = new()
                {
                    InvitationCCRecipient = "dsancbar@emeal.nttdata.com",
                    InvitationMessage = "Hola, te hemos invitado a este tenant uwu"
                }
            });

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public void SerializeObjectsTest()
    {
        BodyUserInfo objetos = new()
        {
            UserBodyRole = new()
            {
                BodyId = "demomvp-cs-ae",
                Role = "members",
            },
            UserInfo = new()
            {
                FirstName = "David",
                LastName = "Sánchez Barragán",
                JobTitle = "Analista NTT Data",
                Office = "Puertollano",
                PrincipalMail = "test@gmail.com"
            }
        };
        string jsonString = System.Text.Json.JsonSerializer.Serialize(objetos);

        Assert.NotECNTy(jsonString);
    }

    [Fact]
    public async Task GetAllBusinessAreasMinistriesTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var BusinessAreasMinistries = await _managementService.GetAllBusinessAreasMinistries(ctx, userUPN, bodyId);
            var json = JsonConvert.SerializeObject(BusinessAreasMinistries);

            Assert.True(BusinessAreasMinistries.Any());
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task AddBusinessAreasMinistriesTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var BusinessUnit = (await _managementService.GetAllBusinessAreasMinistries(ctx, userUPN, bodyId)).FirstOrDefault();
            BusinessUnit!.Division = "fc223273-568d-49d5-a52c-6970e814ef24";
            BusinessUnit!.EndDate = Convert.ToDateTime("2024-12-31T00:00:00+01:00");
            var value = await _managementService.AddOrUpdateBusinessUnit(ctx, userUPN, bodyId, BusinessUnit);
            Assert.True(value != null);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateBusinessAreasMinistriesTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var value = await _managementService.AddOrUpdateBusinessUnit(ctx, userUPN, bodyId, new BusinessUnit
            {
                BusinessArea = "fa9dd202-5d9e-48b8-a527-e06fa881b829",
                Division = "fc223273-568d-49d5-a52c-6970e814ef24",
                StartDate = DateTime.Now,
                EndDate = DateTime.Now.AddDays(365)
            });
            Assert.True(value != null);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateUserInfoTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/{bodyId}");
            var userInfo = new UserInfo()
            {
                UserPrincipalName = "test_email.com@contoso-test.onmicrosoft.com",
                FirstName = "Test editado",
                LastName = "Testez Editado",
                PrincipalMail = "test@email.com",
                Email = "othermail@test2.com",
                JobTitle = "Editado",
                Office = "Castilla-La Mancha",
                BusinessPhone = "666666666",
                CellPhone = "900099900",
                CompanyName = "NTT Data2",
                Department = "DEX",
                EmployeeType = "Analista"
            };
            await _managementService.UpdateUserInformation(ctx, userUPN, bodyId, userInfo);

            Assert.True(true);
        }
        catch (Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task GetAllTermsFromTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");
            var gH = new GraphHelper(ctx);
            //var resultado = await gH.GetAllTermsFrom(siteId, groupId);
            await gH.UpdateSite(ctx.Uri.IdnHost, "demoev-cs-voto", new Microsoft.Graph.Models.Site()
            {
                Description = "New description",
                DisplayName = "New Display name"
            });
            Assert.True(true);
        }
        catch (System.Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateBodyTest()
    {
        try
        {
            var bodyInfo = new Contoso.Portal.Model.Bodies.ConfigDepartments()
            {
                BodyId = "demomvp-cs-ae",
                NombreDepartment = "Test name",
                Department = "ff0e1394-324a-4a85-ac22-6a82417a5b61",
                TipoDepartment = "63ac1500-07b4-4e25-b4ff-6fbd946389ff",
                Secretaria = "707e0e67-86dc-4194-980c-1bf81cacca24",
                Division = "fc223273-568d-49d5-a52c-6970e814ef24",
                dir = "EA002000317",
                SIA = "303328517",
                Materia = "Materia",
                Abreviatura = "CSADR17",
                Period = "718356c0-da6b-4855-9713-aa5514dd7bed",
                BusinessArea = "fa9dd202-5d9e-48b8-a527-e06fa881b82e",
                Activo = true,
                FechaConstitucion = DateTime.Parse("2024-06-17T02:00:00+02:00"),
                FechaExtincion = DateTime.Parse("2026-05-31T11:00:00+02:00"),
                Observaciones = "CSADR17",
                IdEmailBodies = 29,
                IdActBodies = 25,
                IdCertificateBodies = 33,
                IdAgendaBodies = 27,
                IdAttendanceBodies = 31,
                DiasAprodbacionMinutes = 17,
                Description = "Test description holaaa",
                Intersectorial = true,
                TipoMembresia = null,
                IdentificadorConferencia = null
            };
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");
            var response = await _ConfigDepartmentsService.UpdateBodyInformation(ctx, bodyInfo);
            Assert.True(true);
        }
        catch (System.Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task UpdateEXTERNALBodyTest()
    {
        try
        {
            var taxonomyService = new TaxonomyTranslationService(this.MemoryCache, Log<TaxonomyTranslationService>(), Auth, "es-ES", userUPN, "102a5d61-9d78-42e8-ab11-681331d4d37f");
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");

            //var bodies = (await _managementService.GetEXTERNALBodiesInformation(ctx, userUPN, bodyId)).OrderBy(x => x.ID);
            var originalBody = (await _managementService.GetEXTERNALBodiesInformation(ctx, userUPN, bodyId)).OrderBy(x => x.ID).First();
            var taxonomies = (await taxonomyService.GetLabelsByIdAsync([Guid.Parse(originalBody.Department ?? Guid.ECNTy.ToString()), Guid.Parse(originalBody.BusinessArea ?? Guid.ECNTy.ToString()), Guid.Parse(originalBody.Division ?? Guid.ECNTy.ToString())], "es-ES")).Select(x => new { Id = x.Key.ToString(), Label = x.Value });
            var nombreDepartment = taxonomies.Where(x => x.Id == originalBody.Department).Select(x => x.Label).FirstOrDefault() ?? "";

            // bodies.First().NombreDepartment = nombreDepartment + "_NUEVO";
            //bodies.First().CodigoDepartment = "NUEVOCODIGO";
            // bodies.First().Department = bodies.First().Department;
            // bodies.First().TipoDepartmentEXTERNAL = "e29aa354-a2d3-4f72-bdf0-4ee67a04096b";
            // bodies.First().CodigoDepartment = bodies.First().CodigoDepartment + "_Editado";
            // bodies.First().Observaciones = bodies.First().Observaciones + "_Editado";
            // bodies.First().FechaConstitucion = DateTime.Today;
            // bodies.First().FechaExtincion = DateTime.Today.AddDays(1);
            // bodies.First().Activo = true;
            // bodies.First().BusinessArea = "fa9dd202-5d9e-48b8-a527-e06fa881b823";
            // bodies.First().Abreviatura = bodies.First().Abreviatura + "_Editado";
            // bodies.First().FechaInscripcion = DateTime.Today.AddDays(-1);
            // bodies.First().SecretariaText = bodies.First().SecretariaText + "_Editado";
            // bodies.First().StatusDepartment = "68718a9d-8781-4bdd-8aa3-db7598c2e00d";
            // bodies.First().IdCertificateBody = 0;
            // bodies.First().Division = "6d67fdbb-97a0-4fce-8a2a-afeff2fd56b1";

            originalBody.CodigoDepartment = "AAA20120003";

            //var response = await _ConfigDepartmentsService.UpdateEXTERNALBodyInformation(ctx, bodies.First());
            var responseRevert = await _ConfigDepartmentsService.UpdateEXTERNALBodyInformation(ctx, originalBody);
            Assert.True(responseRevert != null);
            //Assert.True(response != null && responseRevert != null);
        }
        catch (System.Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }

    [Fact]
    public async Task CreateEXTERNALBodyTest()
    {
        try
        {
            using var ctx = await CreatePnPContextAsSystem($"sites/Contoso");

            var body = new ConfigEXTERNALBodies
            {
                NombreDepartment = "Nuevo NombreDepartment1",
                Department = null,
                // TipoDepartmentEXTERNAL = "e29aa354-a2d3-4f72-bdf0-4ee67a04096b",
                Inscrito = false,
                // DepartmentAdscripcion = 
                FechaCreacion = DateTime.Today,
                CodigoDepartment = "Nuevo CodigoDepartment",
                Observaciones = "Nuevo Observaciones",
                FechaConstitucion = DateTime.Today,
                FechaExtincion = DateTime.Today.AddDays(1),
                Activo = false,
                BusinessArea = "fa9dd202-5d9e-48b8-a527-e06fa881b823",
                Abreviatura = "Nuevo Abreviatura",
                FechaInscripcion = DateTime.Today.AddDays(-1),
                SecretariaText = "Nuevo SecretariaText",
                // StatusDepartment = "68718a9d-8781-4bdd-8aa3-db7598c2e00d",
                IdCertificateBody = 0,
                Division = "6d67fdbb-97a0-4fce-8a2a-afeff2fd56b1"
            };

            var response = await _managementService.CreateEXTERNALBody(ctx, userUPN, bodyId, body);
            Assert.NotNull(response);
        }
        catch (System.Exception ex)
        {
            Assert.Fail(ex.Message);
        }
    }
}
