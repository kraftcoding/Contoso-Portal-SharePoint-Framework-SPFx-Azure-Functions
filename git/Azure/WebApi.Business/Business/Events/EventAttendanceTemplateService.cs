using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.profile;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.Events
{
    public class EventAttendanceTemplateService(BodyRoleService roleSrv, EventsService eventsService, EventAttendanceService eventAttendanceService,
                                            ConfigDepartmentsService ConfigDepartmentsService, TaxonomyTranslationService taxonomyService, ProfileService ProfileService,
                                            ILogger<EventAttendanceTemplateService> logger, M365AuthHelper auth, DocumentToPdfService pdfService)
        : ServiceBasePnP<EventAttendanceTemplateService>(logger, auth)
    {
        private readonly BodyRoleService _roleSrv = roleSrv;
        private readonly EventsService _eventsService = eventsService;
        private readonly EventAttendanceService _eventAttendanceService = eventAttendanceService;
        private readonly ConfigDepartmentsService _ConfigDepartmentsService = ConfigDepartmentsService;
        private readonly TaxonomyTranslationService _taxonomyService = taxonomyService;
        private readonly ProfileService _ProfileService = ProfileService;
        private readonly DocumentToPdfService _pdfService = pdfService;

        public async Task<byte[]> DownloadAttendaceTemplate(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            try
            {
                (var eventObj, var attendancesObj, var attendancesSharePointObj) = await GetAttendancesInformationByEvent(ctx, bodyId, sharedEventId);
                attendancesObj = attendancesObj != null ? attendancesObj.Where(x => attendancesSharePointObj.ContainsKey(x.UserPrincipalName)).ToArray() : attendancesObj;
                var atendancesXML = await InformationToXML.RetrieveXMLAsync(_taxonomyService, eventObj, attendancesObj, attendancesSharePointObj);
                //var xml = atendancesXML.ToString();
                using var templateStream = await _ConfigDepartmentsService.GetAttendanceTemplate(bodyId);
                var template = await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, atendancesXML);
                return await _pdfService.ConvertToPdf(ctx, template);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(DownloadAttendaceTemplate)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
            }
        }

        private async Task<(Event? eventRet, EventUserAttendance[]? attendancesRet, Dictionary<string, profileInfo> attendancesSharePointRet)>
                GetAttendancesInformationByEvent(IPnPContext ctx, string bodyId, string sharedEventId)
        {

            try
            {
                var eventObj = await _eventsService.GetBySharedEventId(ctx, bodyId, sharedEventId);

                var eventAttendancesObj = await _eventAttendanceService.GetEventAttendance(ctx, bodyId, sharedEventId); // back information for event users
                var usersRetrieveInfoSharePoint = eventAttendancesObj.Where(att => !string.IsNullOrWhiteSpace(att.UserPrincipalName))
                                                                        .Select(att => att.UserPrincipalName)
                                                                        .ToList();
                var eventAttendancesSharePointObj = await _ProfileService.Getprofiles(ctx, usersRetrieveInfoSharePoint); // sharepoint information for event users

                return (eventRet: eventObj,
                        attendancesRet: eventAttendancesObj.ToArray(),
                        attendancesSharePointRet: eventAttendancesSharePointObj);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(GetAttendancesInformationByEvent)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
            }

        }
    }
}
