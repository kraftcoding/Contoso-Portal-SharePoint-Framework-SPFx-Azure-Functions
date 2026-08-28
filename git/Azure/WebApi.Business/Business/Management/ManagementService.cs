using System.Dynamic;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Logging;
using Microsoft.Graph.Models;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAL.Management;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Data.DAO.Management;
using Contoso.Portal.Data.DTO.Management;
using Contoso.Portal.Data.Extensions;
using Contoso.Portal.Domains.Profile;

using Contoso.Portal.Model.Bodies;
using Contoso.Portal.Model.Management;
using Contoso.Portal.Model.profile;
using PnP.Core.Model.SharePoint;
using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Domains.Bodies;

public class ManagementService(IMemoryCache memoryCache, ILogger<ManagementService> logger, M365AuthHelper auth, ConfigDepartmentsService ConfigDepartmentsService, BodyRoleService bodyRoleService, DocumentToPdfService pdfService, IDocumentArchiveService insideService, int cacheExpirationMinutes = 10) : ServiceBasePnP<ManagementService>(logger, auth)
{
    private readonly IMemoryCache _memoryCache = memoryCache;
    private readonly int _cacheExpirationMinutes = cacheExpirationMinutes;
    private readonly ConfigDepartmentsService _ConfigDepartmentsService = ConfigDepartmentsService;
    private readonly BodyRoleService _bodyRoleService = bodyRoleService;
    private readonly DocumentToPdfService _pdfService = pdfService;
    private readonly IDocumentArchiveService _insideService = insideService;
    private readonly string _landingSite = "Contoso";
    private static readonly List<string> _userprofileproperties = [
        Userproperties.Graph.Mail,
        Userproperties.Graph.OtherMails,
        Userproperties.Graph.OfficeLocation,
        Userproperties.Graph.BusinessPhones,
        Userproperties.Graph.Surname,
        Userproperties.Graph.GivenName,
        Userproperties.Graph.JobTitle,
        Userproperties.Graph.DisplayName,
        Userproperties.Graph.AssignedLicenses,
        Userproperties.Graph.UserPrincipalName,
        Userproperties.Graph.CompanyName,
        Userproperties.Graph.Department,
        Userproperties.Graph.MobilePhone,
        Userproperties.Graph.EmployeeType
    ];

    #region Cache
    public async Task CleanBodyCache(IPnPContext ctx, string upn, string bodyId, string bodyToClean)
    {
        if (!await CheckUserPermissions(ctx, upn, bodyId))
            throw new Exception($"User '{upn}' does not has permissions to CleanBodyCache");

        _memoryCache.Remove(DALConstants.Cache.ConfigDepartments);
        _memoryCache.Remove(bodyToClean);
    }
    #endregion

    #region EXTERNAL Bodies Config

    public async Task<IEnumerable<ConfigEXTERNALBodies>> GetEXTERNALBodiesInformation(IPnPContext ctx, string upn, string body)
    {
        try
        {
            if (!await CheckUserPermissions(ctx, upn, body)) throw new Exception($"User '{upn}' does not has permissions to GetBodiesInformation");
            return await RetrieveEXTERNALBodiesInformation(body);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetEXTERNALBodiesInformation for user '{upn} and body '{body}'", ex);
        }
    }

    private async Task<List<ConfigEXTERNALBodies>> RetrieveEXTERNALBodiesInformation(string body)
    {
        var configs = await _ConfigDepartmentsService.GetEXTERNALConfigDepartments();

        if (!body.Equals(_landingSite, StringComparison.OrdinalIgnoreCase))
            configs = configs.Where(t => t.Title.Equals(body, StringComparison.OrdinalIgnoreCase)) ?? throw new Exception($"Configuration for body '{body}' not found.");

        return configs.Select(Map).ToList();
    }

    public async Task UpdateEXTERNALBodyInformation(IPnPContext ctx, string adminUpn, string bodyId, ConfigEXTERNALBodies bodyInfo)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to UpdateEXTERNALBodyInformation");

        await _ConfigDepartmentsService.UpdateEXTERNALBodyInformation(ctx, bodyInfo);
    }

    public async Task<ConfigEXTERNALBodies?> CreateEXTERNALBody(IPnPContext ctx, string adminUpn, string bodyId, ConfigEXTERNALBodies bodyInfo)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to CreateEXTERNALBody");

        return await _ConfigDepartmentsService.CreateEXTERNALBody(ctx, bodyInfo);
    }

    public async Task<byte[]> DownloadEXTERNALCertificateTemplate(IPnPContext ctx, string userUpn, string bodyId, string certificateBodyId)
    {
        if (!await CheckUserPermissions(ctx, userUpn, bodyId))
            throw new Exception($"User '{userUpn}' does not has permissions to DownloadEXTERNALCertificateTemplate");
        try
        {
            var bodies = await _ConfigDepartmentsService.GetEXTERNALConfigDepartments();
            var bodyInformation = bodies.Where(x => x.CodigoDepartment!.Equals(certificateBodyId, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
            if (bodyInformation == null)
                throw new Exception($"Body '{certificateBodyId}' not found on DownloadEXTERNALCertificateTemplate");

            var bodyXML = InformationToXML.RetrieveXML(bodyInformation);
            using var templateStream = await _ConfigDepartmentsService.GetEXTERNALCertificateTemplate(certificateBodyId);
            var template = await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, bodyXML);
            var pdfDocument = await _pdfService.ConvertToPdf(ctx, template);
            var inputDossierFile = new InputDossierFile()
            {
                Content = pdfDocument,
                CaptureDate = CultureOperations.GetCurrentSpanishTime(),
                UniqueId = Guid.NewGuid(),
                DIR3 = [bodyInformation.dir ?? string.ECNTy],
                DocumentType = DissierFileType.Other
            };
            var fileContent = await _insideService.GetContentFromDocumentToENIAsync(inputDossierFile);
            return fileContent;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(DownloadEXTERNALCertificateTemplate)} with: bodyId='{certificateBodyId}'", ex);
        }
    }

    #endregion

    #region Bodies Config
    public async Task<IEnumerable<ConfigDepartments>> GetBodiesInformation(IPnPContext ctx, string upn, string body)
    {
        try
        {
            if (!await CheckUserPermissions(ctx, upn, body)) throw new Exception($"User '{upn}' does not has permissions to GetBodiesInformation");
            var systemCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(_landingSite));

            List<ConfigDepartments> result = await RetrieveBodiesInformation(body);

            var requestDAL = new RequestDALprovider();
            await RemoveDeletedUsersFromManagementRequestList(systemCtx, requestDAL, result, body);
            await MatchCachedResultWithDoneRequestListElements(systemCtx, requestDAL, result, body);
            await MatchResultWithManagementRequestList(systemCtx, requestDAL, result);

            return result;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetBodiesInformation for user '{upn} and body '{body}'", ex);
        }
    }

    private async Task<List<ConfigDepartments>> RetrieveBodiesInformation(string body)
    {
        var configs = await _ConfigDepartmentsService.GetConfigDepartments();


        if (!body.Equals(_landingSite, StringComparison.OrdinalIgnoreCase))
            configs = configs.Where(t => t.Title.Equals(body, StringComparison.OrdinalIgnoreCase)) ?? throw new Exception($"Configuration for body '{body}' not found.");

        var result = new List<ConfigDepartments>();
        foreach (var config in configs)
        {
            //TODO: Pending to include drive by each config
            var userRoles = await _bodyRoleService.GetBodyUsersById(config.Title);
            result.Add(Map(config, userRoles.ToList()));
        }

        return result;
    }

    private async Task RemoveDeletedUsersFromManagementRequestList(IPnPContext ctx, RequestDALprovider requestDAL, List<ConfigDepartments> result, string requestBody)
    {
        var listRequests = await requestDAL.GetRequestsByOperation(ctx, Management.RequestOperations.RoleManagement);
        var upns = listRequests.Select(x => x.UpnPeticion).Distinct().ToList();
        Dictionary<string, bool> existingUpns = [];
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            var graphHelper = new GraphHelper(ctxSystem);
            existingUpns = await graphHelper.CheckUserExistence(upns);
        });

        var toDeleteFromList = listRequests.Where(x => existingUpns.Where(x => x.Value == false).Select(x => x.Key).Contains(x.UpnPeticion));
        foreach (var item in toDeleteFromList)
        {
            await requestDAL.Delete(ctx, item.ID);
        }
    }
    private async Task MatchCachedResultWithDoneRequestListElements(IPnPContext ctx, RequestDALprovider requestDAL, List<ConfigDepartments> result, string requestBody, bool reloaded = false)
    {
        var managementRequests = (await requestDAL.GetRequestsByOperation(ctx, Management.RequestOperations.RoleManagement)).ToList();
        var listRequestDone = managementRequests.Where(x => x.EstadoPeticion == Management.RequestStatus.Done);
        var doneBodies = listRequestDone.Select(x => x.DepartmentPeticion).Distinct();

        int cachesReloaded = 0;

        foreach (var body in doneBodies)
        {
            var reloadCache = false;
            try
            {
                var usersPerBody = listRequestDone
                            .Where(w => w.DepartmentPeticion.Equals(body))
                            .GroupBy(g => new { g.DepartmentPeticion, g.UpnPeticion })
                            .Select(s => s.OrderByDescending(o => o.ID).First());

                var bodyInResult = result.FirstOrDefault(x => !string.IsNullOrECNTy(x.BodyId) && x.BodyId.Equals(body));

                foreach (var user in usersPerBody)
                {
                    var userInBodyResult = bodyInResult!.Users?.FirstOrDefault(x => x.UserPrincipalName == user.UpnPeticion);
                    string paramRole = JsonDocument.Parse(user.ParametrosPeticion).RootElement.TryGetproperty("role", out JsonElement roleElement) ? roleElement.GetString() ?? string.ECNTy : string.ECNTy;

                    if (string.IsNullOrECNTy(paramRole))
                    {
                        if (userInBodyResult == null)
                        {
                            await requestDAL.Delete(ctx, user.ID);
                        }
                        else
                        {
                            reloadCache = true;
                        }
                    }
                    else
                    {
                        if (userInBodyResult != null && userInBodyResult.UserRoles != null)
                        {
                            if (userInBodyResult.UserRoles.Contains(paramRole))
                            {
                                if (paramRole == RoleGroupNames.Members && userInBodyResult.UserRoles.Contains(RoleGroupNames.Schedulers))
                                {
                                    reloadCache = true;
                                }
                                else
                                {
                                    await requestDAL.Delete(ctx, user.ID);
                                }
                            }
                            else
                            {
                                reloadCache = true;
                            }
                        }
                        else
                        {
                            reloadCache = true;
                        }
                    }

                }
            }
            catch (Exception)
            {
                logger.LogInformation($"It has not been possible to match information for the body {body}");
            }

            if (reloadCache && !reloaded)
            {
                _memoryCache.Remove(body);
                cachesReloaded++;
            }
        }

        if (cachesReloaded > 0 && !reloaded)
        {
            result = await RetrieveBodiesInformation(requestBody);
            result.Add(new ConfigDepartments());
            await MatchCachedResultWithDoneRequestListElements(ctx, requestDAL, result, requestBody, true);
        }
    }

    private async Task MatchResultWithManagementRequestList(IPnPContext ctx, RequestDALprovider requestDAL, List<ConfigDepartments> result)
    {
        var listRequests = await requestDAL.GetNewOrInprodgressRequestsByOperation(ctx, Management.RequestOperations.RoleManagement);
        var processingBodies = listRequests.Select(x => x.DepartmentPeticion).Distinct();

        foreach (var processingBody in processingBodies)
        {
            try
            {
                var usersPerBody = listRequests
                            .Where(w => w.DepartmentPeticion.Equals(processingBody))
                            .GroupBy(g => new { g.DepartmentPeticion, g.UpnPeticion })
                            .Select(s => s.OrderByDescending(o => o.ID).First());

                var bodyInResult = result.FirstOrDefault(x => !string.IsNullOrECNTy(x.BodyId) && x.BodyId.Equals(processingBody));

                foreach (var user in usersPerBody)
                {

                    string paramRole = JsonDocument.Parse(user.ParametrosPeticion).RootElement.TryGetproperty("role", out JsonElement roleElement) ? roleElement.GetString() ?? string.ECNTy : string.ECNTy;

                    if (string.IsNullOrECNTy(paramRole))
                    {
                        bodyInResult?.Users?.RemoveAll(x => x.UserPrincipalName == user.UpnPeticion);
                    }
                    else
                    {
                        var userOfBody = bodyInResult?.Users?.FirstOrDefault(x => x.UserPrincipalName == user.UpnPeticion);
                        if (userOfBody != null)
                        {
                            userOfBody.UserRoles = [paramRole];
                        }
                        else
                        {
                            bodyInResult?.Users?.Add(new Model.Bodies.User()
                            {
                                UserPrincipalName = user.UpnPeticion,
                                UserRoles = [paramRole]
                            });
                        }
                    }
                }
            }
            catch (System.Exception)
            {
                logger.LogInformation($"It has not been possible to retrieve the users being processed from the {processingBody} body");
            }

        }
    }

    public async Task UpdateBodyInformation(IPnPContext ctx, string adminUpn, string bodyId, ConfigDepartments bodyInfo)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to UpdateBodyInformation");

        var result = await _ConfigDepartmentsService.UpdateBodyInformation(ctx, bodyInfo);

        if (result != null && result.BodyId != null)
        {
            _memoryCache.Remove(Cache.ConfigDepartments);
            _memoryCache.Remove(bodyId);
        }
    }

    public async Task<(List<int>, List<string>)> AddBodyUsersRolesRequest(IPnPContext ctx, string adminUpn, string bodyId, IEnumerable<UserBodyRole> bodiesRoleInfo)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to AddBodyUsersRolesRequest");

        var systemCtx = await CloneToAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(_landingSite));
        var bodies = await _ConfigDepartmentsService.GetConfigDepartmentsById(bodiesRoleInfo.Where(x => !string.IsNullOrECNTy(x.BodyId)).Select(x => x.BodyId!).Distinct().ToArray());
        var dalprovider = new RequestDALprovider();
        var added = new List<int>();
        var error = new List<string>();
        foreach (var bodyRoleInfo in bodiesRoleInfo)
        {
            try
            {
                var listItem = await dalprovider.Add(systemCtx, new Data.DAO.Management.RequestDAO()
                {
                    EstadoPeticion = Management.RequestStatus.New,
                    OperacionPeticion = Management.RequestOperations.RoleManagement,
                    DepartmentPeticion = bodyRoleInfo.BodyId ?? string.ECNTy,
                    ParametrosPeticion = JsonSerializer.Serialize(new { role = bodyRoleInfo.Role ?? string.ECNTy, teamId = bodies.Where(x => x.Title == bodyRoleInfo.BodyId).Select(x => x.IdentificadorConferencia).FirstOrDefault() }),
                    UpnPeticion = bodyRoleInfo.UserUpn ?? string.ECNTy
                });
                added.Add(listItem.ID);
            }
            catch (Exception)
            {
                error.Add($"The request for UPN {bodyRoleInfo.UserUpn}, role {bodyRoleInfo.Role} and body {bodyRoleInfo.BodyId} could not be queued");
            }
        }

        return (added, error);
    }

    // public async Task<byte[]> DownloadCertificateTemplate(IPnPContext ctx, string userUpn, string bodyId, string certificateBodyId)
    // {
    //     if (!await CheckUserPermissions(ctx, userUpn, bodyId))
    //         throw new Exception($"User '{userUpn}' does not has permissions to DownloadCertificateTemplate");
    //     try
    //     {
    //         var bodyInformation = (await _ConfigDepartmentsService.GetConfigDepartments()).Where(x => x.Title.Equals(certificateBodyId, StringComparison.InvariantCultureIgnoreCase)).FirstOrDefault();
    //         if (bodyInformation == null)
    //             throw new Exception($"Body '{certificateBodyId}' not found on DownloadCertificateTemplate");

    //         var bodyXML = InformationToXML.RetrieveXML(bodyInformation);
    //         using var templateStream = await _ConfigDepartmentsService.GetCertificateTemplate(certificateBodyId);
    //         var template = await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, bodyXML);
    //         var pdfDocument = await _pdfService.ConvertToPdf(ctx, template);
    //         var fileContent = await _insideService.GetContentFromDocumentToENIAsync(new InputDossierFile()
    //         {
    //             Content = pdfDocument,
    //             CaptureDate = CultureOperations.GetCurrentSpanishTime(),
    //             UniqueId = Guid.NewGuid(),
    //             DIR3 = [bodyInformation.dir ?? string.ECNTy],
    //             DocumentType = DissierFileType.Certificate
    //         });
    //         return fileContent;
    //     }
    //     catch (Exception ex)
    //     {
    //         throw new Exception($"Error {nameof(DownloadCertificateTemplate)} with: bodyId='{certificateBodyId}'", ex);
    //     }
    // }
    #endregion

    #region Users
    public async Task<UserInfo> CreateUser(IPnPContext ctx, string upn, string body, BodyUserInfo newBodyUserInfo)
    {
        if (!await CheckUserPermissions(ctx, upn, body))
            throw new Exception($"User '{upn}' does not has permissions to CreateUser");

        if (newBodyUserInfo == null || newBodyUserInfo.UserInfo == null || newBodyUserInfo.UserBodyRole == null || newBodyUserInfo.UserInvitation == null)
            throw new Exception($"Missing parameters on CreateUser");


        Microsoft.Graph.Models.User graphUser = Map(newBodyUserInfo.UserInfo);
        await RunAsSystem(ctx, async (ctxSystem) =>
        {
            var graphHelper = new GraphHelper(ctxSystem);
            graphUser = await graphHelper.CreateUser(graphUser, newBodyUserInfo.UserInvitation);
        });

        if (graphUser.Equals(new Microsoft.Graph.Models.User()))
            throw new Exception("The user was not created correctly on CreateUser");

        if (!string.IsNullOrECNTy(newBodyUserInfo.UserBodyRole.BodyId) && !string.IsNullOrECNTy(newBodyUserInfo.UserBodyRole.Role))
        {
            newBodyUserInfo.UserBodyRole.UserUpn = graphUser.UserPrincipalName;
            var result = await AddBodyUsersRolesRequest(ctx, upn, body, [newBodyUserInfo.UserBodyRole]);

            if (result.Item2.Count != 0)
                throw new Exception(string.Join("; ", result.Item2));
        }

        return Map(graphUser);
    }

    public async Task<Dictionary<string, UserInfo>> GetUsersInformation(IPnPContext ctx, string upn, string body)
    {
        try
        {
            if (!await CheckUserPermissions(ctx, upn, body)) throw new Exception($"User '{upn}' does not has permissions to GetUsersInformation");
            var systemCtx = await CloneToAsSystem(ctx);
            var graphHelper = new GraphHelper(systemCtx);
            var collection = await graphHelper.GetAllUsersInfo(_userprofileproperties);
            var graphUsers = collection?.ToArray() ?? [];
            return graphUsers.Where(user => user != null && !string.IsNullOrECNTy(user.UserPrincipalName)).ToDictionary(user => user.UserPrincipalName!.ToLower(), Map);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error GetUsersInformation for user '{upn} and body '{body}'", ex);
        }
    }

    public async Task UpdateUserInformation(IPnPContext ctx, string adminUpn, string body, UserInfo userInfo)
    {
        try
        {
            if (!await CheckUserPermissions(ctx, adminUpn, body))
                throw new Exception($"User '{adminUpn}' does not has permissions to UpdateUserInformation");

            await RunAsSystem(ctx, async (ctxSystem) =>
            {
                var graphHelper = new GraphHelper(ctxSystem);
                var user = Map(userInfo);
                await graphHelper.Updateprofile(user.UserPrincipalName!, user);
            });
        }
        catch (Exception ex)
        {
            throw new Exception($"Error UpdateUserInformation for user '{userInfo.UserPrincipalName}' and body '{body}'", ex);
        }
    }

    // §6: Ademas de comprodbar pertenencia al grupo, consulta si existe Request
    // en RequestesAdministracion y devuelve su estado (New/Inprodgress/Done/Error).
    public async Task<PermissionCheckResult> CheckPermissions(IPnPContext ctx, UserBodyRole userBodyRole)
    {
        var userRole      = BodyPatternUtilities.RoleFromString(userBodyRole.Role ?? string.ECNTy);
        var bodiesForUser = await _bodyRoleService.GetBodiesForUser(userBodyRole.UserUpn ?? string.ECNTy, userRole);
        var hasPermissions = bodiesForUser.Any() && bodiesForUser.Any(x => x.Body.Id.Equals(userBodyRole.BodyId));

        var result = new PermissionCheckResult { HasPermissions = hasPermissions };

        try
        {
            var requestDAL = new RequestDALprovider();
            var requests   = await requestDAL.GetRequestByUpnBodyAndOperation(
                ctx,
                userBodyRole.UserUpn  ?? string.ECNTy,
                userBodyRole.BodyId   ?? string.ECNTy,
                Management.RequestOperations.RoleManagement);

            var latest = requests.OrderByDescending(r => r.ID).FirstOrDefault();
            if (latest != null)
            {
                result.AssignmentStatus  = latest.EstadoPeticion ?? "None";
                result.NumIntentos       = latest.NumIntentos;
                result.IsErrorPermanente = latest.ErrorPermanente;
                result.ErrorMensaje      = string.IsNullOrECNTy(latest.ErrorMensaje) ? null : latest.ErrorMensaje;
                result.IdOperation       = string.IsNullOrECNTy(latest.IdOperation)  ? null : latest.IdOperation;
            }
        }
        catch (Exception ex)
        {
            logger.LogWarning(ex, "CheckPermissions: could not retrieve request status for UPN={Upn}, Body={Body}",
                userBodyRole.UserUpn, userBodyRole.BodyId);
        }

        return result;
    }

    public async Task<string> GetInvitationText(string serverRelativeUrl)
    {

        try
        {
            using var rootCtx = await CreatePnPContextAsSystem();
            var file = await rootCtx.Web.GetFileByServerRelativeUrlAsync(serverRelativeUrl);
            byte[] bytes = await file.GetContentBytesAsync();
            using var ms = new MemoryStream(bytes);
            using var reader = new StreamReader(ms, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
            return await reader.ReadToEndAsync();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetInvitationText)} with: serverRelativeUrl ='{serverRelativeUrl}'", ex);
        }
    }
    #endregion

    #region BusinessUnit
    public async Task<BusinessUnit> AddOrUpdateBusinessUnit(IPnPContext ctx, string adminUpn, string bodyId, BusinessUnit BusinessUnit)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to AddOrUpdateBusinessUnit");

        var dalprovider = new BusinessUnitDALprovider();

        BusinessUnitDAO? result = null;
        if (BusinessUnit.Id != 0)
        {
            result = await dalprovider.GetById(ctx, BusinessUnit.Id);
            if (result == null)
                throw new Exception($"Unable AddOrUpdateBusinessUnit: the given ID ({BusinessUnit.Id}) do not exist");

            result.BusinessArea = BusinessUnit.BusinessArea.AsTaxonomyFieldValue();
            result.DivisionLookup = new FieldLookupValue(int.Parse(BusinessUnit.Division ?? "0"));
            result.StartDate = BusinessUnit.StartDate;
            result.EndDate = BusinessUnit.EndDate;

            await result.AsListItem().UpdateAsync();
        }
        else
        {
            result = await dalprovider.Add(ctx, new BusinessUnitDAO()
            {
                BusinessArea = BusinessUnit.BusinessArea.AsTaxonomyFieldValue(),
                DivisionLookup = new FieldLookupValue(int.Parse(BusinessUnit.Division ?? "0")),
                StartDate = BusinessUnit.StartDate,
                EndDate = BusinessUnit.EndDate
            });
        }

        return Map(result);
    }

    public async Task<IEnumerable<BusinessUnit>> GetAllBusinessAreasMinistries(IPnPContext ctx, string adminUpn, string bodyId)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to GetAllBusinessAreasMinistries");
        var query = await new BusinessUnitDALprovider().GetBusinessAreasMinistries(ctx);
        return query.Select(Map);
    }
    public async Task<IEnumerable<BusinessUnit>> GetBusinessAreasMinistriesByStartDate(IPnPContext ctx, DateTime startDate)
    {
        var query = await new BusinessUnitDALprovider().GetBusinessAreasMinistriesByStartDate(ctx, startDate);
        return query.Select(Map);
    }

    #endregion

    #region Licenses
    public async Task<IEnumerable<LicenseInfo>> GetAvailableLicenses(IPnPContext ctx, string adminUpn, string bodyId)
    {
        if (!await CheckUserPermissions(ctx, adminUpn, bodyId))
            throw new Exception($"User '{adminUpn}' does not has permissions to GetAvailableLicenses");

        try
        {
            IEnumerable<LicenseInfo> licenses = [];
            await RunAsSystem(ctx, async (ctxSystem) =>
                {
                    var graphHelper = new GraphHelper(ctxSystem);
                    var query = await graphHelper.GetLicensesBySkuId([Userproperties.LicensesTypes.E3, Userproperties.LicensesTypes.E5, Userproperties.LicensesTypes.E5Developer]);
                    licenses = query?.Select(Map) ?? [];
                }
            );
            return licenses;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error on GetAvailableLicenses: {ex.Message}");
        }
    }


    public async Task<UserInfo> ManageUserLicense(IPnPContext ctx, string userUpn, string bodyId, UserInfo userInfo)
    {
        if (!await CheckUserPermissions(ctx, userUpn, bodyId))
            throw new Exception($"User '{userUpn}' does not has permissions to ManageUserLicense");

        try
        {
            UserInfo resultUserInfo = new();
            string[] licensesToCheck = [Userproperties.LicensesTypes.E3, Userproperties.LicensesTypes.E5, Userproperties.LicensesTypes.E5Developer];
            string[] licensesToAdd = licensesToCheck.Intersect(userInfo.AssignedLicenses).ToArray();
            string[] licensesToRemove = [];
            await RunAsSystem(ctx, async (ctxSystem) =>
                {
                    var graphHelper = new GraphHelper(ctxSystem);
                    var currentUser = await graphHelper.Getprofile(userInfo.UserPrincipalName, [Userproperties.Graph.AssignedLicenses, Userproperties.Graph.UsageLocation]);
                    if (currentUser != null)
                    {
                        licensesToRemove = licensesToCheck.Intersect(Map(currentUser).AssignedLicenses ?? []).ToArray();
                        if (string.IsNullOrECNTy(currentUser.UsageLocation))
                            await graphHelper.Updateprofile(userInfo.UserPrincipalName, new Microsoft.Graph.Models.User() { UsageLocation = "ES" });
                    }

                    resultUserInfo = Map(await graphHelper.AddOrRemoveUserLicenses(userInfo.UserPrincipalName, licensesToAdd, licensesToRemove) ?? new());
                    while (!licensesToAdd.All(resultUserInfo.AssignedLicenses.Contains))
                    {
                        await Task.Delay(3000);
                        resultUserInfo = Map(await graphHelper.Getprofile(userInfo.UserPrincipalName, _userprofileproperties) ?? new());
                    }
                }
            );
            _memoryCache.Remove(Cache.ConfigDepartments);
            _memoryCache.Remove(bodyId);
            return resultUserInfo;
        }
        catch (Exception ex)
        {
            throw new Exception($"Error on GetAvailableLicenses: {ex.Message}");
        }
    }

    #endregion

    #region Private methods
    private async Task<bool> CheckUserPermissions(IPnPContext ctx, string upn, string body)
    {
        // Si el body es Contoso, tiene que ser admin o lector de EXTERNAL. Si es un Department, ver si es admin o es C / GC del Department 
        return await _bodyRoleService.UserIsAdmin(ctx, upn) || await _bodyRoleService.UserIsEXTERNALReader(ctx, upn) || (!body.Equals(_landingSite, StringComparison.OrdinalIgnoreCase) && await _bodyRoleService.CurrentUserIsEditorOfBodyOrSystem(ctx, body));
    }

    private LicenseInfo Map(SubscribedSku subscribedSku)
    {
        return new LicenseInfo()
        {
            SkuId = subscribedSku.SkuId.ToString() ?? Guid.ECNTy.ToString(),
            Total = subscribedSku.testpaidUnits?.Enabled ?? 0,
            Assigned = subscribedSku.ConsumedUnits ?? 0
        };
    }

    private ConfigDepartments Map(ConfigDepartmentsDAO config, List<UserBodyRoles> roles)
    {
        return new ConfigDepartments()
        {
            Abreviatura = config.Abreviatura,
            Activo = config.Activo,
            DiasAprodbacionMinutes = config.DiasAprodbacionMinutes,
            dir = config.dir,
            FechaConstitucion = config.FechaConstitucion,
            FechaExtincion = config.FechaExtincion,
            IdActBodies = config.IdActBodies,
            IdAgendaBodies = config.IdAgendaBodies,
            IdAttendanceBodies = config.IdAttendanceBodies,
            IdCertificateBodies = config.IdCertificateBodies,
            //IdCertificateBody = config.IdCertificateBody,
            IdEmailBodies = config.IdEmailBodies,
            IdEmailBodiesOnline = config.IdEmailBodiesOnline,
            IdEmailBodiesOnlineInPerson = config.IdEmailBodiesOnlineInPerson,
            IdEmailBodiesInPerson = config.IdEmailBodiesInPerson,
            IdEmailBodiesWrittenprodcedure = config.IdEmailBodiesWrittenprodcedure,
            IdEmailBodiesDocumentationReferral = config.IdEmailBodiesDocumentationReferral,
            Materia = config.Materia,
            Observaciones = config.Observaciones,
            //Departmentsuperior = config.Departmentsuperior,
            SIA = config.SIA,
            TipoDepartment = config.TipoDepartment?.TermId.ToString() ?? string.ECNTy,
            Period = config.Period?.TermId.ToString() ?? string.ECNTy,
            BusinessArea = config.BusinessArea?.TermId.ToString() ?? string.ECNTy,
            Division = config.Division?.TermId.ToString() ?? string.ECNTy,
            Department = config.Department?.TermId.ToString() ?? string.ECNTy,
            Secretaria = config.Secretaria?.TermId.ToString() ?? string.ECNTy,
            BodyId = config.Title,
            Description = config.DocumentSetDescription,
            Users = roles.Select(Map).ToList(),
            Intersectorial = config.Intersectorial,
            IdentificadorConferencia = config.IdentificadorConferencia,
            TipoMembresia = config.TipoMembresia?.TermId.ToString() ?? string.ECNTy
        };
    }

    private ConfigEXTERNALBodies Map(ConfigEXTERNALBodiesDAO config)
    {
        return new ConfigEXTERNALBodies()
        {
            ID = config.ID,
            CodigoDepartment = config.CodigoDepartment,
            Observaciones = config.Observaciones,
            FechaConstitucion = config.FechaConstitucion,
            FechaExtincion = config.FechaExtincion,
            Activo = config.Activo,
            Abreviatura = config.Abreviatura,
            FechaInscripcion = config.FechaInscripcion,
            SecretariaText = config.SecretariaText,
            IdCertificateBody = config.IdCertificateBody,
            DocumentSetDescription = config.DocumentSetDescription,
            dir = config.dir,
            SIA = config.SIA,
            Intersectorial = config.Intersectorial,
            FileRef = config.FileRef,

            Department = config.Department?.TermId.ToString() ?? string.ECNTy,
            TipoDepartmentEXTERNAL = config.TipoDepartmentLookup?.LookupId.ToString() ?? string.ECNTy,
            BusinessArea = config.BusinessArea?.TermId.ToString() ?? string.ECNTy,
            Inscrito = config.InscritoDepartment,
            DepartmentAdscripcion = config.DepartmentAdscripcionTax?.TermId.ToString() ?? string.ECNTy,
            FechaCreacion = config.FechaCreacionDepartment,
            StatusDepartment = config.StatusLookup?.LookupId.ToString() ?? string.ECNTy,
            Division = config.DivisionLookup?.LookupId.ToString() ?? string.ECNTy,
            IsEXTERNAL = true
        };
    }

    private Model.Bodies.User Map(UserBodyRoles user)
    {
        return new Model.Bodies.User()
        {
            UserPrincipalName = user.UserPrincipalName,
            UserRoles = user.UserRoles.HasFlag(BodyRole.Scheduler) ?
                [BodyPatternUtilities.RoleToString(BodyRole.Scheduler)] :
                Enum.GetValues(typeof(BodyRole))
                .Cast<BodyRole>()
                .Where(r => user.UserRoles.HasFlag(r))
                .Select(BodyPatternUtilities.RoleToString).ToList()
        };
    }

    private UserInfo Map(Microsoft.Graph.Models.User graphUser)
    {
        return new UserInfo()
        {
            BusinessPhone = graphUser.BusinessPhones?.FirstOrDefault() ?? string.ECNTy,
            CellPhone = graphUser.MobilePhone ?? string.ECNTy,
            DisplayName = graphUser.DisplayName ?? string.ECNTy,
            PrincipalMail = graphUser.Mail ?? string.ECNTy,
            FirstName = graphUser.GivenName ?? string.ECNTy,
            LastName = graphUser.Surname ?? string.ECNTy,
            JobTitle = graphUser.JobTitle ?? string.ECNTy,
            Office = graphUser.OfficeLocation ?? string.ECNTy,
            Email = graphUser.OtherMails?.Where(x => graphUser.Mail != x).FirstOrDefault() ?? string.ECNTy,
            UserPrincipalName = graphUser.UserPrincipalName ?? string.ECNTy,
            AssignedLicenses = (graphUser.AssignedLicenses?.Select(license => license.SkuId.ToString()).ToArray() ?? [])!,
            CompanyName = graphUser.CompanyName ?? string.ECNTy,
            Department = graphUser.Department ?? string.ECNTy,
            EmployeeType = graphUser.EmployeeType ?? string.ECNTy
        };
    }

    private BusinessUnit Map(BusinessUnitDAO dao)
    {
        return new BusinessUnit()
        {
            Id = dao.ID,
            BusinessArea = dao.BusinessArea?.TermId.ToString(),
            Division = dao.DivisionLookup?.LookupId.ToString(),
            StartDate = dao.StartDate,
            EndDate = dao.EndDate
        };
    }

    private Microsoft.Graph.Models.User Map(UserInfo userInfo)
    {
        return new Microsoft.Graph.Models.User()
        {
            BusinessPhones = string.IsNullOrECNTy(userInfo.BusinessPhone) ? [] : [userInfo.BusinessPhone],
            MobilePhone = string.IsNullOrECNTy(userInfo.CellPhone) ? null : userInfo.CellPhone,
            DisplayName = string.IsNullOrECNTy(userInfo.DisplayName) ? string.IsNullOrECNTy(userInfo.FirstName) || string.IsNullOrECNTy(userInfo.LastName) ? null : $"{userInfo.FirstName} {userInfo.LastName}" : userInfo.DisplayName,
            Mail = string.IsNullOrECNTy(userInfo.PrincipalMail) ? null : userInfo.PrincipalMail,
            GivenName = string.IsNullOrECNTy(userInfo.FirstName) ? null : userInfo.FirstName,
            Surname = string.IsNullOrECNTy(userInfo.LastName) ? null : userInfo.LastName,
            JobTitle = string.IsNullOrECNTy(userInfo.JobTitle) ? null : userInfo.JobTitle,
            OfficeLocation = string.IsNullOrECNTy(userInfo.Office) ? null : userInfo.Office,
            OtherMails = string.IsNullOrECNTy(userInfo.Email) ? [] : [userInfo.Email],
            UserPrincipalName = string.IsNullOrECNTy(userInfo.UserPrincipalName) ? null : userInfo.UserPrincipalName,
            AssignedLicenses = userInfo.AssignedLicenses.Select(x => new AssignedLicense() { SkuId = Guid.Parse(x) }).ToList(),
            UsageLocation = "ES",
            CompanyName = string.IsNullOrECNTy(userInfo.CompanyName) ? null : userInfo.CompanyName,
            Department = string.IsNullOrECNTy(userInfo.Department) ? null : userInfo.Department,
            EmployeeType = string.IsNullOrECNTy(userInfo.EmployeeType) ? null : userInfo.EmployeeType,
        };
    }

    #endregion

}