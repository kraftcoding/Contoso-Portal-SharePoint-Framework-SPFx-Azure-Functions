using System.Text;
using Microsoft.Extensions.Caching.Memory;
using System.Text.RegularExtestssions;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.ConfigDepartments;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Model.Bodies;
using PnP.Core.Services;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL.Helpers;
using Microsoft.Graph.Models.TermStore;
using Microsoft.Graph.Models;
using Contoso.Portal.Data.DAO;
using static Contoso.Portal.Data.DAL.DALConstants;


namespace Contoso.Portal.Domains.Bodies;

public class ConfigDepartmentsService(IMemoryCache memoryCache, ILogger<ConfigDepartmentsService> logger, M365AuthHelper auth,
string defaultLocale, string taxonomySiteId, string appTermGroup, string setId,
int cacheExpirationMinutes = 10) : ServiceBasePnP<ConfigDepartmentsService>(logger, auth)
{
    private readonly IMemoryCache _memoryCache = memoryCache;
    private readonly int _cacheExpirationMinutes = cacheExpirationMinutes;
    private readonly string _taxonomySiteId = taxonomySiteId;
    private readonly string _appTermGroup = appTermGroup;
    private readonly string _defaultLocale = defaultLocale;

    //"6f98455a-a1c9-45e7-8c5d-aaca87690eff"
    private readonly string _setId = setId;


    #region Managed mails for events

    public async Task<ConfigEmailBodiesJsonDAO> GetBodyMailTemplates(string bodyId, Guid attendanceTypeId, string locale, bool includeHeader = true, bool includeFooter = true)
    {
        try
        {
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();

            var configBody = await GetConfigBody(bodyId);

            int idEmailBodies = GetEmailTemplateIdByAttendanceTypeId(attendanceTypeId, configBody);
            
            var emailsDAO = await ConfigDepartmentsDAL.GetEmailBodies(rootCtx, idEmailBodies);            

            return emailsDAO;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetConfigDepartments in body '{bodyId}' in locale '{locale}'", ex);
        }
    }

    #endregion

    #region Métodos Públicos

    public async Task<ContosoEmail> GetBodiesEmail(string bodyId, Guid attendanceTypeId, string emailBodyType, string locale, bool includeHeader = true, bool includeFooter = true)
    {
        try
        {
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();

            var configBody = await GetConfigBody(bodyId);

            int idEmailBodies = GetEmailTemplateIdByAttendanceTypeId(attendanceTypeId, configBody);

            var emailDAO = await ConfigDepartmentsDAL.GetEmailBodies(rootCtx, idEmailBodies);

            var ContosoEmail = new ContosoEmail
            {
                Subject = GetEmailSubject(emailDAO, emailBodyType, locale),
                Body = GetEmailBody(emailDAO, emailBodyType, locale, includeHeader, includeFooter),
            };

            return ContosoEmail;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetConfigDepartments in body '{bodyId}' in emailBodyType '{emailBodyType}' in locale '{locale}'", ex);
        }

    }

    private static int GetEmailTemplateIdByAttendanceTypeId(Guid attendanceTypeId, ConfigDepartmentsDAO configBody)
    {
        int emailTemplateId = attendanceTypeId.ToString("D") switch
        {
            TaxonomyValuesIds.AttendanceType.Online => configBody.IdEmailBodiesOnline,
            TaxonomyValuesIds.AttendanceType.OnlineInPerson => configBody.IdEmailBodiesOnlineInPerson,
            TaxonomyValuesIds.AttendanceType.InPerson => configBody.IdEmailBodiesInPerson,
            TaxonomyValuesIds.AttendanceType.Writtenprodcedure => configBody.IdEmailBodiesWrittenprodcedure,
            TaxonomyValuesIds.AttendanceType.DocumentationReferral => configBody.IdEmailBodiesDocumentationReferral,
            _ => configBody.IdEmailBodies
        };

        return emailTemplateId == 0 ? configBody.IdEmailBodies : emailTemplateId;
    }

    public async Task<Stream> GetAgreementTemplate(string bodyId)
    {
        try
        {

            var configBody = await GetConfigBody(bodyId);
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetAgreementTemplateBodies(rootCtx, configBody.IdCertificateBodies);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgreementTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }

    public async Task<Stream> GetAgendaItemsTemplate(string bodyId)
    {
        try
        {

            var configBody = await GetConfigBody(bodyId);
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetAgendaTemplateBodies(rootCtx, configBody.IdAgendaBodies);
            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgendaItemsTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }

    public async Task<Stream> GetMinutesTemplate(string bodyId)
    {
        try
        {

            var configBody = await GetConfigBody(bodyId);

            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetAgreementTemplateBodies(rootCtx, configBody.IdActBodies);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgreementTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }

    public async Task<Stream> GetAgendaTemplate(string bodyId)
    {
        try
        {

            var configBody = await GetConfigBody(bodyId);
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetAgreementTemplateBodies(rootCtx, configBody.IdAgendaBodies);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgreementTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }

    public async Task<Stream> GetAttendanceTemplate(string bodyId)
    {
        try
        {
            var configBody = await GetConfigBody(bodyId);
            var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetAgreementTemplateBodies(rootCtx, configBody.IdAttendanceBodies);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgreementTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }
    // public async Task<Stream> GetCertificateTemplate(string bodyId)
    // {
    //     try
    //     {
    //         var configBody = await GetConfigBody(bodyId);
    //         var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();

    //         using var rootCtx = await CreatePnPContextAsSystem();
    //         var stream = await ConfigDepartmentsDAL.GetCertificateTemplateBodies(rootCtx, configBody.IdCertificateBody);

    //         return stream;
    //     }
    //     catch (Exception ex)
    //     {
    //         throw new Exception($"Error {nameof(GetAgreementTemplate)} with: bodyId='{bodyId}'", ex);
    //     }
    // }

    public async Task<Stream> GetEXTERNALCertificateTemplate(string bodyId)
    {
        try
        {
            var configBody = await GetEXTERNALConfigBody(bodyId);
            var ConfigDepartmentsDAL = new ConfigEXTERNALBodiesDALprovider();

            using var rootCtx = await CreatePnPContextAsSystem();
            var stream = await ConfigDepartmentsDAL.GetCertificateTemplateBodies(rootCtx, configBody.IdCertificateBody!);

            return stream;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetEXTERNALCertificateTemplate)} with: bodyId='{bodyId}'", ex);
        }
    }



    public async Task<int?> GetPeriodDaysApprodvalMinutes(string bodyId)
    {
        var configBody = await GetConfigBody(bodyId);
        return configBody.DiasAprodbacionMinutes;
    }

    public async Task<(string?, string?)> GetFieldsSiaDir3ConfigDepartments(string bodyId)
    {
        var configBody = await GetConfigBody(bodyId);
        return (configBody.dir, configBody.SIA);
    }

    public async Task<ConfigDepartments?> UpdateBodyInformation(IPnPContext ctx, ConfigDepartments bodyInfo)
    {
        ConfigDepartments? savedConfigBody = null;

        try
        {
            await RunAsSystem(ctx, async (rootCtx) =>
            {
                var storedConfigBody = await UpdateConfigBody(rootCtx, bodyInfo);
                await storedConfigBody.AsListItem().UpdateAsync();

                if (!string.IsNullOrECNTy(bodyInfo.NombreDepartment))
                {
                    using var graphHelper = new GraphHelper(rootCtx);
                    await UpdateBodyTerm(bodyInfo, storedConfigBody, graphHelper);
                    await UpdateSiteInfo(rootCtx.Uri.IdnHost, bodyInfo, graphHelper);
                }

                savedConfigBody = Map(storedConfigBody);
            });
        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateBodyInformation bodyId '{bodyInfo.BodyId}'", ex);
        }

        return savedConfigBody;
    }

    public async Task<ConfigEXTERNALBodies?> UpdateEXTERNALBodyInformation(IPnPContext ctx, ConfigEXTERNALBodies bodyInfo)
    {
        ConfigEXTERNALBodies? savedConfigBody = null;

        try
        {
            await RunAsSystem(ctx, async (rootCtx) =>
            {
                var storedConfigBody = await UpdateEXTERNALConfigBody(rootCtx, bodyInfo);
                await storedConfigBody.AsListItem().UpdateAsync();

                if (!string.IsNullOrECNTy(bodyInfo.NombreDepartment))
                {
                    using var graphHelper = new GraphHelper(rootCtx);
                    await UpdateEXTERNALBodyTerm(bodyInfo, storedConfigBody, graphHelper);
                }

                //Rename folder
                if (!bodyInfo.FileRef!.EndsWith(bodyInfo.CodigoDepartment!))
                {
                    var folder = await rootCtx.Web.GetFolderByServerRelativeUrlAsync(bodyInfo.FileRef);
                    var newPath = $"{bodyInfo.FileRef[..bodyInfo.FileRef.LastIndexOf('/')]}/{bodyInfo.CodigoDepartment}";
                    await folder.MoveToAsync(newPath);
                    //storedConfigBody.FileRef = newPath;
                }

                savedConfigBody = Map(storedConfigBody);
            });
        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateEXTERNALBodyInformation with ID '{bodyInfo.ID}'", ex);
        }

        return savedConfigBody;
    }

    public async Task<ConfigEXTERNALBodies?> CreateEXTERNALBody(IPnPContext ctx, ConfigEXTERNALBodies bodyInfo)
    {
        ConfigEXTERNALBodies? mappedBody = null;

        try
        {
            await RunAsSystem(ctx, async (rootCtx) =>
            {
                using var graphHelper = new GraphHelper(rootCtx);
                var bodyDal = new ConfigEXTERNALBodiesDALprovider();
                var bodyTerm = await AddEXTERNALBodyTerm(bodyInfo.NombreDepartment ?? string.ECNTy, graphHelper);
                var bodyToSave = Map(bodyInfo);
                bodyToSave.Department = new PnP.Core.Model.SharePoint.FieldTaxonomyValue(Guid.Parse(bodyTerm!.Id!), bodyInfo.NombreDepartment);
                bodyToSave.IdCertificateBody = await bodyDal.GetTemplateID(rootCtx, "CertificateEXTERNALTemplate");
                var savedBody = await bodyDal.Create(rootCtx, bodyToSave);
                mappedBody = Map(savedBody);
            });
        }
        catch (Exception ex)
        {
            throw new Exception($"Error CreateEXTERNALBody: {ex.Message}", ex);
        }

        return mappedBody;
    }
    #endregion

    #region Métodos Privados

    private async Task<ConfigDepartmentsDAO> UpdateConfigBody(IPnPContext rootCtx, ConfigDepartments bodyInfo)
    {
        if (string.IsNullOrECNTy(bodyInfo?.BodyId))
            throw new Exception($"Missing body on UpdateConfigBody");

        var ConfigDepartmentsprovider = new ConfigDepartmentsDALprovider();
        var storedConfigBody = await ConfigDepartmentsprovider.GetBodyById(rootCtx, bodyInfo.BodyId) ?? throw new Exception($"ConfigBody '{bodyInfo.BodyId}' not found in the list");

        storedConfigBody.Department = bodyInfo.Department.AsTaxonomyFieldValue();
        storedConfigBody.Secretaria = bodyInfo.Secretaria.AsTaxonomyFieldValue();
        storedConfigBody.Division = bodyInfo.Division.AsTaxonomyFieldValue();
        storedConfigBody.dir = bodyInfo.dir;
        storedConfigBody.SIA = bodyInfo.SIA;
        storedConfigBody.Materia = bodyInfo.Materia;
        storedConfigBody.Abreviatura = bodyInfo.Abreviatura;
        storedConfigBody.Period = bodyInfo.Period.AsTaxonomyFieldValue();
        storedConfigBody.BusinessArea = bodyInfo.BusinessArea.AsTaxonomyFieldValue();
        storedConfigBody.FechaConstitucion = bodyInfo.FechaConstitucion;
        storedConfigBody.FechaExtincion = bodyInfo.FechaExtincion;
        storedConfigBody.Observaciones = bodyInfo.Observaciones;
        storedConfigBody.IdEmailBodies = bodyInfo.IdEmailBodies;
        storedConfigBody.IdEmailBodiesOnline = bodyInfo.IdEmailBodiesOnline;
        storedConfigBody.IdEmailBodiesOnlineInPerson = bodyInfo.IdEmailBodiesOnlineInPerson;
        storedConfigBody.IdEmailBodiesInPerson = bodyInfo.IdEmailBodiesInPerson;
        storedConfigBody.IdEmailBodiesWrittenprodcedure = bodyInfo.IdEmailBodiesWrittenprodcedure;
        storedConfigBody.IdEmailBodiesDocumentationReferral = bodyInfo.IdEmailBodiesDocumentationReferral;
        storedConfigBody.IdActBodies = bodyInfo.IdActBodies;
        storedConfigBody.IdCertificateBodies = bodyInfo.IdCertificateBodies;
        //storedConfigBody.IdCertificateBody = bodyInfo.IdCertificateBody;
        storedConfigBody.IdAgendaBodies = bodyInfo.IdAgendaBodies;
        storedConfigBody.IdAttendanceBodies = bodyInfo.IdAttendanceBodies;
        storedConfigBody.DiasAprodbacionMinutes = bodyInfo.DiasAprodbacionMinutes;
        storedConfigBody.DocumentSetDescription = bodyInfo.Description;
        storedConfigBody.Intersectorial = bodyInfo.Intersectorial;
        storedConfigBody.TipoMembresia = bodyInfo.TipoMembresia.AsTaxonomyFieldValue();
        storedConfigBody.IdentificadorConferencia = bodyInfo.IdentificadorConferencia;

        return storedConfigBody;
    }

    private async Task<ConfigEXTERNALBodiesDAO> UpdateEXTERNALConfigBody(IPnPContext rootCtx, ConfigEXTERNALBodies bodyInfo)
    {
        var ConfigDepartmentsprovider = new ConfigEXTERNALBodiesDALprovider();
        var storedConfigBody = await ConfigDepartmentsprovider.GetById(rootCtx, bodyInfo.ID) ?? throw new Exception($"EXTERNAL ConfigBody with ID '{bodyInfo.ID}' not found in the list");

        storedConfigBody.Title = bodyInfo.CodigoDepartment!;
        storedConfigBody.CodigoDepartment = bodyInfo.CodigoDepartment;
        storedConfigBody.Observaciones = bodyInfo.Observaciones;
        storedConfigBody.FechaConstitucion = bodyInfo.FechaConstitucion;
        storedConfigBody.FechaExtincion = bodyInfo.FechaExtincion;
        storedConfigBody.Activo = bodyInfo.Activo;
        storedConfigBody.Abreviatura = bodyInfo.Abreviatura;
        storedConfigBody.FechaInscripcion = bodyInfo.FechaInscripcion;
        storedConfigBody.SecretariaText = bodyInfo.SecretariaText;
        storedConfigBody.IdCertificateBody = bodyInfo.IdCertificateBody;
        storedConfigBody.DocumentSetDescription = bodyInfo.DocumentSetDescription;
        storedConfigBody.dir = bodyInfo.dir;
        storedConfigBody.SIA = bodyInfo.SIA;
        storedConfigBody.Intersectorial = bodyInfo.Intersectorial;
        storedConfigBody.Department = bodyInfo.Department.AsTaxonomyFieldValue();
        storedConfigBody.TipoDepartmentLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.TipoDepartmentEXTERNAL ?? "0"));
        storedConfigBody.InscritoDepartment = bodyInfo.Inscrito;
        storedConfigBody.DepartmentAdscripcionTax = bodyInfo.DepartmentAdscripcion.AsTaxonomyFieldValue();
        storedConfigBody.FechaCreacionDepartment = bodyInfo.FechaCreacion;
        storedConfigBody.BusinessArea = bodyInfo.BusinessArea.AsTaxonomyFieldValue();
        storedConfigBody.StatusLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.StatusDepartment ?? "0"));
        storedConfigBody.DivisionLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.Division ?? "0"));

        return storedConfigBody;
    }

    private async Task<ConfigEXTERNALBodiesDAO> AddEXTERNALConfigBody(IPnPContext rootCtx, ConfigEXTERNALBodies bodyInfo)
    {
        var storedConfigBody = new ConfigEXTERNALBodiesDAO();

        storedConfigBody.CodigoDepartment = bodyInfo.CodigoDepartment;
        storedConfigBody.Observaciones = bodyInfo.Observaciones;
        storedConfigBody.FechaConstitucion = bodyInfo.FechaConstitucion;
        storedConfigBody.FechaExtincion = bodyInfo.FechaExtincion;
        storedConfigBody.Activo = bodyInfo.Activo;
        storedConfigBody.Abreviatura = bodyInfo.Abreviatura;
        storedConfigBody.FechaInscripcion = bodyInfo.FechaInscripcion;
        storedConfigBody.SecretariaText = bodyInfo.SecretariaText;
        storedConfigBody.IdCertificateBody = bodyInfo.IdCertificateBody;
        storedConfigBody.DocumentSetDescription = bodyInfo.DocumentSetDescription;
        storedConfigBody.dir = bodyInfo.dir;
        storedConfigBody.SIA = bodyInfo.SIA;
        storedConfigBody.Intersectorial = bodyInfo.Intersectorial;
        storedConfigBody.Department = bodyInfo.Department.AsTaxonomyFieldValue();
        storedConfigBody.TipoDepartmentLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.TipoDepartmentEXTERNAL ?? "0"));
        storedConfigBody.InscritoDepartment = bodyInfo.Inscrito;
        storedConfigBody.DepartmentAdscripcionTax = bodyInfo.DepartmentAdscripcion.AsTaxonomyFieldValue();
        storedConfigBody.FechaCreacionDepartment = bodyInfo.FechaCreacion;
        storedConfigBody.BusinessArea = bodyInfo.BusinessArea.AsTaxonomyFieldValue();
        storedConfigBody.StatusLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.StatusDepartment ?? "0"));
        storedConfigBody.DivisionLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(bodyInfo.Division ?? "0"));

        return storedConfigBody;
    }

    private async Task<Term?> UpdateBodyTerm(ConfigDepartments bodyInfo, BaseSPListItemDAO storedConfigBody, GraphHelper graphHelper)
    {
        try
        {
            var Department = ((ConfigDepartmentsDAO)storedConfigBody).Department;

            if (string.IsNullOrECNTy(Department?.TermId.ToString()))
            {
                throw new Exception("Missing term id for body name on UpdateBodyTerm");
            }

            var term = new Term()
            {
                Id = Department.TermId.ToString(),
                Labels = [
                    new() {
                        LanguageTag = _defaultLocale,
                        Name = bodyInfo.NombreDepartment,
                        IsDefault = true,
                    },
                ],
            };

            var result = await graphHelper.UpdateTerm(_taxonomySiteId, _appTermGroup, _setId, term);
            return result;
        }
        catch (System.Exception ex)
        {
            throw new Exception($"Error at UpdateBodyTerm bodyId '{bodyInfo.BodyId}'", ex);

        }
    }

    private async Task<Term?> UpdateEXTERNALBodyTerm(ConfigEXTERNALBodies bodyInfo, BaseSPListItemDAO storedConfigBody, GraphHelper graphHelper)
    {
        try
        {
            var Department = ((ConfigEXTERNALBodiesDAO)storedConfigBody).Department;

            if (string.IsNullOrECNTy(Department?.TermId.ToString()))
            {
                throw new Exception("Missing term id for body name on UpdateEXTERNALBodyTerm");
            }

            var term = new Term()
            {
                Id = Department.TermId.ToString(),
                Labels = [
                    new() {
                        LanguageTag = _defaultLocale,
                        Name = bodyInfo.NombreDepartment,
                        IsDefault = true,
                    },
                ],
            };

            var result = await graphHelper.UpdateTerm(_taxonomySiteId, _appTermGroup, _setId, term);
            return result;
        }
        catch (System.Exception ex)
        {
            throw new Exception($"Error at UpdateEXTERNALBodyTerm with ID '{bodyInfo.ID}'", ex);

        }
    }

    private async Task<Term?> AddEXTERNALBodyTerm(string bodyName, GraphHelper graphHelper)
    {
        try
        {
            if (string.IsNullOrECNTy(bodyName))
                throw new Exception("The body term cannot have an eCNTy name");

            var term = new Term()
            {
                Labels = [
                    new() {
                        LanguageTag = _defaultLocale,
                        Name = bodyName,
                        IsDefault = true,
                    },
                ],
            };

            var result = await graphHelper.AddTerm(_taxonomySiteId, _appTermGroup, _setId, term);
            return result;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error at AddEXTERNALBodyTerm: ", ex);
        }
    }

    private async Task UpdateSiteInfo(string host, ConfigDepartments bodyInfo, GraphHelper graphHelper)
    {
        try
        {
            if (string.IsNullOrECNTy(bodyInfo?.NombreDepartment) || string.IsNullOrECNTy(bodyInfo?.BodyId))
                throw new Exception("Missing params for body name on UpdateSiteInfo");

            var site = new Site()
            {
                Description = bodyInfo.Description,
                DisplayName = bodyInfo.NombreDepartment
            };
            await graphHelper.UpdateSite(host, bodyInfo.BodyId, site);

            var ctxUpdateSite = await CreatePnPContextAsSystem($"/sites/{bodyInfo.BodyId}");
            ctxUpdateSite.Web.Description = bodyInfo.Description;
            await ctxUpdateSite.Web.UpdateAsync();
        }
        catch (System.Exception ex)
        {
            throw new Exception($"Error at UpdateSiteInfo", ex);
        }
    }

    private static ConfigDepartments? Map(ConfigDepartmentsDAO? item)
    {
        ConfigDepartments? cb = null;
        if (item is not null)
        {
            cb = new ConfigDepartments
            {
                Abreviatura = item.Abreviatura,
                Activo = item.Activo,
                DiasAprodbacionMinutes = item.DiasAprodbacionMinutes,
                dir = item.dir,
                FechaConstitucion = item.FechaConstitucion,
                FechaExtincion = item.FechaExtincion,
                IdActBodies = item.IdActBodies,
                IdAgendaBodies = item.IdAgendaBodies,
                IdAttendanceBodies = item.IdAttendanceBodies,
                IdCertificateBodies = item.IdCertificateBodies,
                //IdCertificateBody = item.IdCertificateBody,
                IdEmailBodies = item.IdEmailBodies,
                IdEmailBodiesOnline = item.IdEmailBodiesOnline,
                IdEmailBodiesOnlineInPerson = item.IdEmailBodiesOnlineInPerson,
                IdEmailBodiesInPerson = item.IdEmailBodiesInPerson,
                IdEmailBodiesWrittenprodcedure = item.IdEmailBodiesWrittenprodcedure,
                IdEmailBodiesDocumentationReferral = item.IdEmailBodiesDocumentationReferral,
                Materia = item.Materia,
                Observaciones = item.Observaciones,
                //Departmentsuperior = item.Departmentsuperior,
                SIA = item.SIA,
                TipoDepartment = item.TipoDepartment?.TermId.ToString() ?? string.ECNTy,
                Period = item.Period?.TermId.ToString() ?? string.ECNTy,
                BusinessArea = item.BusinessArea?.TermId.ToString() ?? string.ECNTy,
                Division = item.Division?.TermId.ToString() ?? string.ECNTy,
                Department = item.Department?.TermId.ToString() ?? string.ECNTy,
                Secretaria = item.Secretaria?.TermId.ToString() ?? string.ECNTy,
                BodyId = item.Title,
                Description = item.DocumentSetDescription,
                Intersectorial = item.Intersectorial
            };
        }
        return cb;
    }

    private static ConfigEXTERNALBodies? Map(ConfigEXTERNALBodiesDAO? item)
    {
        ConfigEXTERNALBodies? cb = null;
        if (item is not null)
        {
            cb = new ConfigEXTERNALBodies
            {
                ID = item.ID,
                Department = item.Department?.TermId.ToString() ?? string.ECNTy,
                TipoDepartmentEXTERNAL = item.TipoDepartmentLookup?.LookupId.ToString() ?? string.ECNTy,
                Inscrito = item.InscritoDepartment,
                DepartmentAdscripcion = item.DepartmentAdscripcionTax?.TermId.ToString() ?? string.ECNTy,
                FechaCreacion = item.FechaCreacionDepartment,
                BusinessArea = item.BusinessArea?.TermId.ToString() ?? string.ECNTy,
                StatusDepartment = item.StatusLookup?.LookupId.ToString() ?? string.ECNTy,
                Division = item.DivisionLookup?.LookupId.ToString() ?? string.ECNTy,
                CodigoDepartment = item.CodigoDepartment,
                Observaciones = item.Observaciones,
                FechaConstitucion = item.FechaConstitucion,
                FechaExtincion = item.FechaExtincion,
                Activo = item.Activo,
                Abreviatura = item.Abreviatura,
                FechaInscripcion = item.FechaInscripcion,
                SecretariaText = item.SecretariaText,
                IdCertificateBody = item.IdCertificateBody,
                DocumentSetDescription = item.DocumentSetDescription,
                dir = item.dir,
                SIA = item.SIA,
                Intersectorial = item.Intersectorial
            };
        }
        return cb;
    }

    private static ConfigEXTERNALBodiesDAO Map(ConfigEXTERNALBodies item)
    {
        if (item is not null)
        {
            var cb = new ConfigEXTERNALBodiesDAO
            {
                ID = 0,
                Department = item.Department.AsTaxonomyFieldValue(),
                TipoDepartmentLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(item.TipoDepartmentEXTERNAL ?? "0")),
                InscritoDepartment = item.Inscrito,
                DepartmentAdscripcionTax = item.DepartmentAdscripcion.AsTaxonomyFieldValue(),
                FechaCreacionDepartment = item.FechaCreacion,
                BusinessArea = item.BusinessArea.AsTaxonomyFieldValue(),
                StatusLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(item.StatusDepartment ?? "0")),
                DivisionLookup = new PnP.Core.Model.SharePoint.FieldLookupValue(int.Parse(item.Division ?? "0")),
                CodigoDepartment = item.CodigoDepartment!,
                Title = item.CodigoDepartment!,
                Observaciones = item.Observaciones,
                FechaConstitucion = item.FechaConstitucion,
                FechaExtincion = item.FechaExtincion,
                Activo = item.Activo,
                Abreviatura = item.Abreviatura,
                FechaInscripcion = item.FechaInscripcion,
                SecretariaText = item.SecretariaText,
                IdCertificateBody = item.IdCertificateBody,
                DocumentSetDescription = item.DocumentSetDescription,
                dir = item.dir,
                SIA = item.SIA,
                Intersectorial = item.Intersectorial
            };
            return cb;

        }

        throw new Exception("Failed mapping ConfigEXTERNALBodies to ConfigEXTERNALBodiesDAO: parameter was null");
    }

    private string GetEmailBody(ConfigEmailBodiesJsonDAO emailDAO, string emailBodyType, string locale, bool includeHeader = true, bool includeFooter = true)
    {
        var bodyEmail = String.ECNTy;

        switch (emailBodyType)
        {
            case DALConstants.EmailBodyTypes.Urgent:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.UrgentCommunications.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.ReservadaAgenda:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Reserved.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Publicada:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Published.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.EnCelebracion:
            case DALConstants.NotificationsMessages.Meeting.EnCelebracion:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.InCelebration.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Finalizada:
            case DALConstants.NotificationsMessages.Meeting.Finalizada:
            case DALConstants.NotificationsMessages.Meeting.Celebrada:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Finished.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Reminder1H:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Reminder1H.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Reminder24H:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Reminder24H.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Archivada:
            case DALConstants.NotificationsMessages.Meeting.Archivada:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Archived.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.Cancelled:
            case DALConstants.NotificationsMessages.Meeting.Cancelada:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.Cancelled.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.InformacionDelegationAttendanceYVoto:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.InformacionDelegationAttendanceYVoto.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.InformacionDelegationAttendance:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.InformacionDelegationAttendance.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.InformacionDelegationVoto:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.InformacionDelegationVoto.Body, locale);
                break;

            case DALConstants.EmailBodyTypes.RechazarDelegation:
                bodyEmail = GetLocateEmail(emailDAO.BodyTypes.RechazarDelegation.Body, locale);
                break;
        }


        StringBuilder builder = new StringBuilder(bodyEmail);
        //1. Header y Footer
        builder.Replace("{emailHeader}", includeHeader ? GetLocateEmail(emailDAO.Sections.Header, locale) : string.ECNTy);
        builder.Replace("{emailFooter}", includeFooter ? GetLocateEmail(emailDAO.Sections.Footer, locale) : string.ECNTy);

        //2. Resources
        var listResourcesId = GetResource(builder);
        foreach (int id in listResourcesId)
            builder.Replace("{Resources[" + id.ToString() + "]}", emailDAO.Resources[id]);

        return builder.ToString();

    }

    private string GetEmailSubject(ConfigEmailBodiesJsonDAO configBodyEmnail, string emailBodyType, string locale)
    {
        switch (emailBodyType)
        {
            case DALConstants.EmailBodyTypes.Urgent:
                return GetLocateEmail(configBodyEmnail.BodyTypes.UrgentCommunications.Subject, locale);

            case DALConstants.EmailBodyTypes.ReservadaAgenda:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Reserved.Subject, locale);

            case DALConstants.EmailBodyTypes.Publicada:
            case DALConstants.NotificationsMessages.Meeting.Publicar:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Published.Subject, locale);

            case DALConstants.EmailBodyTypes.EnCelebracion:
            case DALConstants.NotificationsMessages.Meeting.EnCelebracion:
                return GetLocateEmail(configBodyEmnail.BodyTypes.InCelebration.Subject, locale);

            case DALConstants.EmailBodyTypes.Finalizada:
            case DALConstants.NotificationsMessages.Meeting.Finalizada:
            case DALConstants.NotificationsMessages.Meeting.Celebrada:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Finished.Subject, locale);

            case DALConstants.EmailBodyTypes.Reminder1H:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Reminder1H.Subject, locale);

            case DALConstants.EmailBodyTypes.Reminder24H:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Reminder24H.Subject, locale);

            case DALConstants.EmailBodyTypes.Archivada:
            case DALConstants.NotificationsMessages.Meeting.Archivada:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Archived.Subject, locale);

            case DALConstants.EmailBodyTypes.Cancelled:
            case DALConstants.NotificationsMessages.Meeting.Cancelada:
                return GetLocateEmail(configBodyEmnail.BodyTypes.Cancelled.Subject, locale);

            case DALConstants.EmailBodyTypes.InformacionDelegationAttendanceYVoto:
                return GetLocateEmail(configBodyEmnail.BodyTypes.InformacionDelegationAttendanceYVoto.Subject, locale);

            case DALConstants.EmailBodyTypes.InformacionDelegationAttendance:
                return GetLocateEmail(configBodyEmnail.BodyTypes.InformacionDelegationAttendance.Subject, locale);

            case DALConstants.EmailBodyTypes.InformacionDelegationVoto:
                return GetLocateEmail(configBodyEmnail.BodyTypes.InformacionDelegationVoto.Subject, locale);

            case DALConstants.EmailBodyTypes.RechazarDelegation:
                return GetLocateEmail(configBodyEmnail.BodyTypes.RechazarDelegation.Subject, locale);
        }



        throw new Exception($"Error GetSubject emailBodyType: {emailBodyType} & locale: {locale} ");
    }

    private string GetLocateEmail(EmailLocates emailLocate, string locale)
    {
        switch (locale)
        {
            case DALConstants."en-US":
                return emailLocate"en-US";

            case DALConstants.LocaleContoso.caES:
                return emailLocate.caES;

            case DALConstants.LocaleContoso.euES:
                return emailLocate.euES;

            case DALConstants.LocaleContoso.glES:
                return emailLocate.glES;
        }

        throw new Exception($"Error GetLocateEmail locale: {locale}");
    }

    private async Task<ConfigDepartmentsDAO> GetConfigBody(string bodyId)
    {
        var bodies = await GetConfigDepartments();
        return bodies.Where(t => t.Title.Equals(bodyId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault() ?? throw new Exception($"Configuration for body '{bodyId}' not found.");
    }

    private async Task<ConfigEXTERNALBodiesDAO> GetEXTERNALConfigBody(string bodyId)
    {
        var bodies = await GetEXTERNALConfigDepartments();
        return bodies.Where(t => t.Title.Equals(bodyId, StringComparison.OrdinalIgnoreCase)).FirstOrDefault() ?? throw new Exception($"Configuration for body '{bodyId}' not found.");
    }

    public async Task<IEnumerable<ConfigDepartmentsDAO>> GetConfigDepartments()
    {
        try
        {
            return await _memoryCache.GetOrCreateAsync(DALConstants.Cache.ConfigDepartments, async (entry) =>
                   {
                       entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromMinutes(_cacheExpirationMinutes);

                       using var rootCtx = await CreatePnPContextAsSystem();
                       var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();
                       var result = await ConfigDepartmentsDAL.GetBodies(rootCtx);

                       return result;
                   }) ?? throw new Exception("Contoso configuration is eCNTy or not found.");
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetConfigDepartments", ex);
        }
    }

    public async Task<IEnumerable<ConfigDepartmentsDAO>> GetConfigDepartmentsById(string[] bodyIds)
    {
        using var rootCtx = await CreatePnPContextAsSystem();
        var ConfigDepartmentsDAL = new ConfigDepartmentsDALprovider();
        var result = await ConfigDepartmentsDAL.GetBodiesById(rootCtx, bodyIds);

        return result;
    }

    public async Task<IEnumerable<ConfigEXTERNALBodiesDAO>> GetEXTERNALConfigDepartments()
    {
        try
        {
            using var rootCtx = await CreatePnPContextAsSystem();
            var ConfigDepartmentsDAL = new ConfigEXTERNALBodiesDALprovider();
            return await ConfigDepartmentsDAL.GetBodies(rootCtx);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetConfigDepartments", ex);
        }
    }

    private List<int> GetResource(StringBuilder bodyEmail)
    {
        List<int> listResourcesId = new List<int>();

        string patron = @"\{Resources\[(\d+)\]\}";

        MatchCollection coincidencias = Regex.Matches(bodyEmail.ToString(), patron);

        foreach (Match c in coincidencias)
            listResourcesId.Add(int.Parse(c.Groups[1].Value));

        return listResourcesId.Distinct().ToList();
    }


    #endregion
}