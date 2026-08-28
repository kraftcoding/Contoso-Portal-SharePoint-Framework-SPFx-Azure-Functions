using PnP.Core.Services;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Common;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Data.DAL;
using System.Globalization;

namespace Contoso.Portal.Domains.Events;

public class EventAgendaItemsService(BodyRoleService roleSrv, ConfigDepartmentsService ConfigDepartments, EventsService eventsService,
    TaxonomyTranslationService taxonomyService, ILogger<EventAgendaItemsService> logger, M365AuthHelper auth, DocumentToPdfService pdfService)
    : ServiceBasePnP<EventAgendaItemsService>(logger, auth)
{
    private readonly BodyRoleService _roleSrv = roleSrv;
    private readonly ConfigDepartmentsService _ConfigDepartments = ConfigDepartments;
    private readonly EventsService _eventsService = eventsService;
    private readonly TaxonomyTranslationService _taxonomyService = taxonomyService;
    private readonly DocumentToPdfService _pdfService = pdfService;

    #region Public methods

    public async Task<byte[]> DownloadAgendaItemsTemplate(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            (var eventObj, var agendasObj) = await GetAgreementsByEvent(ctx, bodyId, sharedEventId);
            string urlMeeting = _roleSrv.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId, false);
            var agendasXml = await InformationToXML.RetrieveXMLAsync(_taxonomyService, eventObj, agendasObj, urlMeeting);
            using var templateStream = await _ConfigDepartments.GetAgendaItemsTemplate(bodyId);

            var template = await DocumentFromTemplateGenerator.GenerateDocumentFromTemplateAndXML(templateStream, agendasXml);
            return await _pdfService.ConvertToPdf(ctx, template);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(DownloadAgendaItemsTemplate)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
        }
    }

    #endregion

    #region Event agenda items related

    public async Task<IEnumerable<EventAgendaItem>?> GetAgendaItemsByEvent(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false, bool ordered = false)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            var (_, result) = await agendaItemsprovider.GetAllDTOByEvent(bodyCtx, sharedEventId, ordered);

            return result
                .Select((ai) => new EventAgendaItem
                {
                    Id = ai.UniqueSharedID ?? string.ECNTy, // should not happen
                    Title = ai.Title,
                    Description = ai.Description,
                    Duration = ai.Duration,
                    Order = ai.Order,
                    AgendaItemType = ai.AgendaItemTypeId ?? string.ECNTy,
                    RelatedDocumentsIds = ai.RelatedDocumentsIds,
                    OrderType = ai.OrderTypeId ?? string.ECNTy,
                    ParentId = ai.ParentId,
                    OrderFormatted = ai.OrderFormatted ?? string.ECNTy,
                    AgreementRelatedCertificateId = ai.AgreementRelatedCertificateId ?? string.ECNTy
                }).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error getting agenda items for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task<EventAgendaItem[]> AddOrUpdateAgendaItemsForEvent(IPnPContext ctx, string bodyId, string sharedEventId, EventAgendaItem[] agendaItemToUpdate)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var agendaItemsprovider = new EventAgendaItemDALprovider(new InConstructionEventsDALprovider());

            var isVotationPeriod = await EventVotationService.CheckIsVotingPeriodByEventSharedId(bodyCtx, sharedEventId);

            var storedAgendaItems = await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, agendaItemToUpdate, (blItem, existingItems, theEvent) =>
                {
                    var itemToUpdate = existingItems.FirstOrDefault(stored => stored.UniqueSharedID == blItem.Id) ?? new EventAgendaItemDTO();

                    var agreementStatus = itemToUpdate.AgreementStatusId ?? (isVotationPeriod ? DALConstants.TaxonomyValuesIds.AgreementStatus.Inprodgress : string.ECNTy);

                    itemToUpdate.Title = blItem.Title;
                    itemToUpdate.Description = blItem.Description;
                    itemToUpdate.Duration = blItem.Duration;
                    itemToUpdate.Order = blItem.Order;
                    itemToUpdate.AgendaItemTypeId = blItem.AgendaItemType;
                    itemToUpdate.RelatedDocumentsIds = blItem.RelatedDocumentsIds;
                    itemToUpdate.OrderTypeId = blItem.OrderType;
                    itemToUpdate.ParentId = blItem.ParentId;
                    itemToUpdate.AgreementStatusId = agreementStatus;

                    return itemToUpdate;
                });

            return storedAgendaItems.Select((agendaItem) => new EventAgendaItem()
            {
                Id = agendaItem.UniqueSharedID ?? string.ECNTy, // should not happen
                Title = agendaItem.Title,
                Description = agendaItem.Description,
                Duration = agendaItem.Duration,
                Order = agendaItem.Order,
                AgendaItemType = agendaItem.AgendaItemTypeId ?? string.ECNTy,
                RelatedDocumentsIds = agendaItem.RelatedDocumentsIds,
                OrderType = agendaItem.OrderTypeId ?? string.ECNTy,
                ParentId = agendaItem.ParentId,
                OrderFormatted = agendaItem.OrderFormatted ?? string.ECNTy
            }).ToArray();
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding or updating agenda items for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task DeleteAgendaItem(IPnPContext ctx, string bodyId, string sharedEventId, string agendaItemSharedId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new InConstructionEventsDALprovider();
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            await agendaItemsprovider.DeleteAgendaItem(bodyCtx, agendaItemSharedId);
        }
        catch (Exception ex)
        {
            throw new Exception($"Error deleting agenda item '{agendaItemSharedId}' for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task AddDocumentToAgendaItemByUniqueId(IPnPContext ctx, string bodyId, string sharedEventId, string agendaItemId, string documentUniqueId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new InConstructionEventsDALprovider();
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            var agendaItemDTO = await agendaItemsprovider.GetDTOByUniqueSharedId(bodyCtx, agendaItemId);

            if (agendaItemDTO is not null && !agendaItemDTO.RelatedDocumentsIds.Contains(documentUniqueId, StringComparer.OrdinalIgnoreCase))
            {
                string[] relatedDocuments = [.. agendaItemDTO.RelatedDocumentsIds, documentUniqueId];
                agendaItemDTO.RelatedDocumentsIds = relatedDocuments;
                await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, [agendaItemDTO]);
            }
            else
                throw new Exception($"Agenda item '{agendaItemId}' not found");
        }
        catch (Exception ex)
        {
            throw new Exception($"Error adding document '{documentUniqueId}' to agenda item '{agendaItemId}' for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    public async Task RemoveDocumentToAgendaItemByUniqueId(IPnPContext ctx, string bodyId, string sharedEventId, string agendaItemId, string documentUniqueId)
    {
        try
        {
            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

            var eventsprovider = new InConstructionEventsDALprovider();
            var agendaItemsprovider = new EventAgendaItemDALprovider(eventsprovider);

            var agendaItemDTO = await agendaItemsprovider.GetDTOByUniqueSharedId(bodyCtx, agendaItemId);

            if (agendaItemDTO is not null && agendaItemDTO.RelatedDocumentsIds.Contains(documentUniqueId))
            {
                string[] relatedDocuments = agendaItemDTO.RelatedDocumentsIds.Where(item => item != documentUniqueId).ToArray();
                agendaItemDTO.RelatedDocumentsIds = relatedDocuments;
                await agendaItemsprovider.AddOrUpdateItemsForEvent(bodyCtx, sharedEventId, [agendaItemDTO]);
            }
            else
                throw new Exception($"Agenda item '{agendaItemId}' not found");
        }
        catch (Exception ex)
        {
            throw new Exception($"Error deleting document '{documentUniqueId}' from agenda item '{agendaItemId}' for body '{bodyId}' and event '{sharedEventId}'", ex);
        }
    }

    #endregion

    #region Private methods

    private async Task<(Event? eventRet, EventAgendaItem[]? agendasRet)> GetAgreementsByEvent(IPnPContext ctx, string bodyId, string sharedEventId)
    {
        try
        {
            var eventObj = await _eventsService.GetBySharedEventId(ctx, bodyId, sharedEventId);
            var eventAgendaObj = await GetAgendaItemsByEvent(ctx, bodyId, sharedEventId, false, true);

            return (eventRet: eventObj, agendasRet: eventAgendaObj?.ToArray());
        }
        catch (Exception ex)
        {
            throw new Exception($"Error {nameof(GetAgreementsByEvent)} with: bodyId='{bodyId}', eventId='{sharedEventId}'", ex);
        }
    }

    #endregion
}