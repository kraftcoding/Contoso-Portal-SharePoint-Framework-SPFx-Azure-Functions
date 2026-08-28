using System.Data;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using PnP.Core.Services;
using Contoso.Portal.Model.profile;
using Contoso.Portal.Domains.Profile;
using Microsoft.Extensions.Logging;

namespace Contoso.Portal.Domains.Events
{
    public class EventMinutesTemplateService : ServiceBasePnP<EventMinutesTemplateService>
    {
        private readonly BodyRoleService _roleSrv;
        private readonly ConfigDepartmentsService _ConfigDepartmentsService;
        private readonly EventsService _eventsService;
        private readonly EventAttendanceService _eventAttendanceService;
        private readonly EventMinutesService _eventMinutesService;
        private readonly TaxonomyTranslationService _taxonomyService;
        private readonly EventVotationService _eventVotationService;
        private readonly EventAgendaItemsService _eventAgendaItemService;
        private readonly EventAgreementService _eventAgreementService;


        public EventMinutesTemplateService(BodyRoleService roleSrv, ConfigDepartmentsService ConfigDepartmentsService, EventsService eventsService
            , EventAttendanceService eventAttendanceService, EventMinutesService eventMinutesService, TaxonomyTranslationService taxonomyService, EventAgendaItemsService eventAgendaItemService, EventAgreementService eventAgreementService, EventVotationService eventVotationService
            , ILogger<EventMinutesTemplateService> logger, M365AuthHelper auth)
            : base(logger, auth)
        {
            _roleSrv = roleSrv;
            _ConfigDepartmentsService = ConfigDepartmentsService;
            _eventsService = eventsService;
            _eventAttendanceService = eventAttendanceService;
            _eventMinutesService = eventMinutesService;
            _taxonomyService = taxonomyService;
            _eventAgendaItemService = eventAgendaItemService;
            _eventAgreementService = eventAgreementService;
            _eventVotationService = eventVotationService;
        }

        public async Task<byte[]> DownloadMinutesTemplate(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            try
            {
                (var eventObj, var agendasObj,
                    var attendancesObj, var attendancesSharePointObj,
                    var agreementsObj, var minutesObj, var summaries) = await GetMinutesInformationByEvent(ctx, bodyId, sharedEventId);
                string rootSiteUrl = _roleSrv.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId, false);
                string rootSiteDocumentUrl = _roleSrv.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId, true);
                var publishedDocumentsDalprovider = new EventPublishedDocumentDALprovider();
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
                var publishedDocuments = await publishedDocumentsDalprovider.GetPublishedDocument(bodyCtx, sharedEventId);
                attendancesObj = attendancesObj != null ? attendancesObj.Where(x => attendancesSharePointObj.ContainsKey(x.UserPrincipalName)).ToArray() : attendancesObj;
                var minutesXml = await InformationToXML.RetrieveXMLAsync(_taxonomyService, eventObj, agendasObj,
                                                                                attendancesObj, attendancesSharePointObj,
                                                                                agreementsObj, minutesObj,
                                                                                rootSiteDocumentUrl, rootSiteUrl, publishedDocuments, summaries);
                //var xml = minutesXml.ToString();
                using var templateStream = await _ConfigDepartmentsService.GetMinutesTemplate(bodyId);
                return await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, minutesXml);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(DownloadMinutesTemplate)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
            }
        }


        #region Private methods

        private async Task<(Event? eventRet, EventAgendaItem[]? agendasRet,
                            EventUserAttendance[]? attendancesRet, Dictionary<string, profileInfo> attendancesSharePointRet,
                            EventAgreement[]? agreementsRet, EventMinutesInformation? minutesRet, Dictionary<string, Summary> summaries)>
            GetMinutesInformationByEvent(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            try
            {
                // bodyObj (not by now)
                var eventObj = await _eventsService.GetBySharedEventId(ctx, bodyId, sharedEventId);

                var eventAgendaObj = await _eventAgendaItemService.GetAgendaItemsByEvent(ctx, bodyId, sharedEventId, false, true);

                var eventAttendancesObj = await _eventAttendanceService.GetEventAttendance(ctx, bodyId, sharedEventId); // back information for event users
                var usersRetrieveInfoSharePoint = eventAttendancesObj
                                                    .Where(att => !string.IsNullOrWhiteSpace(att.UserPrincipalName))
                                                    .Select(att => att.UserPrincipalName)//.ToList() // Users attendance
                                                    .Concat(
                                                        eventAttendancesObj
                                                            .Where(att => !string.IsNullOrWhiteSpace(att.DelegationUserPrincipalName))
                                                            .Select(att => att.DelegationUserPrincipalName)//.ToList() // Users with attendance delegated
                                                        )
                                                    .Concat(
                                                        eventAttendancesObj
                                                            .Where(vote => !string.IsNullOrWhiteSpace(vote.DelegationUserVotePrincipalName))
                                                            .Select(vote => vote.DelegationUserVotePrincipalName)//.ToList() // Users with vote delegated
                                                    )
                                                    .Distinct()
                                                    .ToList();
                var eventAttendancesSharePointObj = await new ProfileService().Getprofiles(ctx, usersRetrieveInfoSharePoint); // sharepoint information for event users

                var eventAgreementsObj = await _eventAgreementService.GetAgreementsByEvent(ctx, bodyId, sharedEventId);

                var eventVotingSummary = await _eventVotationService.GetVotingSummaries(ctx, bodyId, sharedEventId, eventAgreementsObj?.Select(ag => ag.Id));

                var eventMinutesObj = await _eventMinutesService.GetMinutesInformationByEvent(ctx, bodyId, sharedEventId);

                return (eventRet: eventObj,
                        agendasRet: eventAgendaObj?.ToArray(),
                        attendancesRet: eventAttendancesObj.ToArray(),
                        attendancesSharePointRet: eventAttendancesSharePointObj,
                        agreementsRet: eventAgreementsObj.ToArray(),
                        minutesRet: eventMinutesObj,
                        summaries: eventVotingSummary);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(GetMinutesInformationByEvent)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
            }
        }

        #endregion


    }
}
