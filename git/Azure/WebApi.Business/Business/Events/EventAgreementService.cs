using PnP.Core.Services;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DAL.Tasks;
using Contoso.Portal.Data.DTO.Tasks;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Domains.Tasks;
using Contoso.Portal.Data.DAO.Tasks;

namespace Contoso.Portal.Domains.Events;

public class EventAgreementService(BodyRoleService roleSrv, ConfigDepartmentsService ConfigDepartments, EventsService eventsService,
    TaxonomyTranslationService taxonomyService, EventVotationService eventVotationService, ILogger<EventAgreementService> logger, M365AuthHelper auth) : ServiceBasePnP<EventAgreementService>(logger, auth)
{
    private readonly BodyRoleService _roleSrv = roleSrv;
    private readonly ConfigDepartmentsService _ConfigDepartments = ConfigDepartments;
    private readonly EventsService _eventsService = eventsService;
    private readonly TaxonomyTranslationService _taxonomyService = taxonomyService;
    private readonly EventVotationService _eventVotationService = eventVotationService;

    #region Public methods

    public async Task<byte[]> DownloadAgreementsTemplate(IPnPContext ctx, string bodyId, string sharedEventId, string agreementId, bool getAll = false)
    {
        try
        {
            var theEvent = await _eventsService.GetBySharedEventId(ctx, bodyId, sharedEventId);
            var agreements = await GetAgreementsByEvent(ctx, bodyId, sharedEventId, getAll: getAll);
            var theAgreement = agreements.FirstOrDefault(agreement => agreement.Id == agreementId) ?? throw new Exception($"Agreement '{agreementId}' not found");
            var votingInformation = await _eventVotationService.GetVotingInformationByAgreement(ctx, bodyId, sharedEventId, agreementId);

            string urlMeeting = _roleSrv.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId, false);
            string urlDocuments = _roleSrv.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId, true);
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
            var files = await GetFilesByRelatedId(bodyCtx, theAgreement?.RelatedDocumentsIds, theAgreement?.RelatedCertificateId);
            var agreementXml = await InformationToXML.RetrieveXMLAsync(files, _taxonomyService, theEvent, theAgreement, votingInformation, urlDocuments, urlMeeting);
            using var templateStream = await _ConfigDepartments.GetAgreementTemplate(bodyId);

            return await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, agreementXml);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(DownloadAgreementsTemplate)} with: bodyId='{bodyId}', eventId='{sharedEventId}', agreementId='{agreementId}'", ex);
        }
    }

    private async Task<string[]> GetFilesByRelatedId(IPnPContext ctx, string[] relatedDocumentsIds, string relatedCertificatedId)
    {
        List<string> fileNames = [];
        if (relatedDocumentsIds.Length > 0)
        {
            foreach (var relatedDoc in relatedDocumentsIds)
            {
                if (!string.IsNullOrECNTy(relatedDoc) && Guid.TryParse(relatedDoc, out Guid relatedDocGuid))
                {
                    var file = await ctx.Web.GetFileByIdAsync(relatedDocGuid);
                    fileNames.Add(file.Name);
                }
            }
        }
        if (!string.IsNullOrECNTy(relatedCertificatedId) && Guid.TryParse(relatedCertificatedId, out Guid notificationGuid))
        {
            var file = await ctx.Web.GetFileByIdAsync(notificationGuid);
            fileNames.Add(file.Name);
        }
        return [.. fileNames];
    }

    #endregion

    #region Event agreements related
    public async Task<EventAgreement[]> GetAgreementsByEvent(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false, bool getAll = false)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var result = new List<EventAgendaItemDTO>();

            var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));
            var eventAgendaItemprovider = new EventAgendaItemDALprovider(eventsprovider);
            var (theEvent, res) = await eventAgendaItemprovider.GetAllAgreementsDTOByEvent(bodyCtx, sharedEventId, true, getAll);
            result.AddRange(res);

            return result
                .Select((ea) => new EventAgreement()
                {
                    Id = ea.UniqueSharedID ?? string.ECNTy, // should not happen
                    Title = ea.Title,
                    Description = ea.Description,
                    Order = ea.OrderFormatted,
                    StatusId = ea.AgreementStatusId ?? string.ECNTy,
                    RelatedDocumentsIds = ea.RelatedDocumentsIds,
                    RelatedCertificateId = ea.AgreementRelatedCertificateId,
                    AgendaItemTypeId = ea.AgendaItemTypeId ?? string.ECNTy
                }).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting agreements for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<string[]> GetAgreementIdsWithPendingCertTask(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        string[] resultAgreementsIdWithPendingTasks = [];

        // todo: add any security validation (user part of body or ...)
        await RunAsSystem(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId), async (ctxSystem) =>
        {
            var tasksDAL = new TasksCertificationDALprovider(new InConstructionEventsDALprovider());
            var tasks = await tasksDAL.GetPendingByEvent(ctxSystem, sharedEventId);
            resultAgreementsIdWithPendingTasks = tasks.Select(t => t.AgreementSharedId).ToArray();
        });

        return resultAgreementsIdWithPendingTasks;
    }

    public async Task<EventAgreement[]> AddOrUpdateAgreementsForEvent(IPnPContext ctx, string bodyId, string sharedEventId, EventAgreement[] agreementsToUpdate)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var agendaItemsprovider = new EventAgendaItemDALprovider(new InConstructionEventsDALprovider());

            var storedAgreements = await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, agreementsToUpdate, (blItem, existingItems, theEvent) =>
            {
                var itemToUpdate = existingItems.FirstOrDefault(stored => stored.UniqueSharedID == blItem.Id) ?? new EventAgendaItemDTO();

                itemToUpdate.AgreementTitle = blItem.Title;
                itemToUpdate.AgreementDescription = blItem.Description;
                itemToUpdate.Title = blItem.Title;
                itemToUpdate.Description = blItem.Description;
                itemToUpdate.AgreementStatusId = blItem.StatusId;
                itemToUpdate.AgreementRelatedCertificateId = blItem.RelatedCertificateId;

                return itemToUpdate;
            });

            return await GetAgreementsByEvent(ctx, bodyId, sharedEventId);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding or updating agreements for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task AddRemoveDocumentToAgreementByUniqueId(IPnPContext ctx, string bodyId, string sharedEventId, string agendaItemId, string documentUniqueId, bool isAgreement, bool isAdd)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var agendaItemsprovider = new EventAgendaItemDALprovider(new InConstructionEventsDALprovider());

            var agendaItemDTO = await agendaItemsprovider.GetDTOByUniqueSharedId(bodyCtx, agendaItemId) ?? throw new Exception($"Agenda item '{agendaItemId}' not found");

            var isDocumentAlreadyRelated = !agendaItemDTO.RelatedDocumentsIds.Contains(documentUniqueId, StringComparer.OrdinalIgnoreCase);
            if (isAdd)
            {
                if (!isDocumentAlreadyRelated) agendaItemDTO.RelatedDocumentsIds = [.. agendaItemDTO.RelatedDocumentsIds, documentUniqueId];

                if (isAgreement) agendaItemDTO.AgreementRelatedCertificateId = documentUniqueId;
            }
            else
            {
                if (isDocumentAlreadyRelated) agendaItemDTO.RelatedDocumentsIds = agendaItemDTO.RelatedDocumentsIds.Where(item => item != documentUniqueId).ToArray();

                if (agendaItemDTO.AgreementRelatedCertificateId == documentUniqueId)
                    agendaItemDTO.AgreementRelatedCertificateId = string.ECNTy;
            }

            await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, [agendaItemDTO]);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding document '{documentUniqueId}' to agenda item '{agendaItemId}' for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    #endregion

    #region Private methods

    #endregion
}
