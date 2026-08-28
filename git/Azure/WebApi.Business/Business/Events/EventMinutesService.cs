using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Event;
using Contoso.Portal.Data.DTO.Event;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Model.Events;
using PnP.Core.Services;

namespace Contoso.Portal.Domains.Events
{
    public class EventMinutesService : ServiceBasePnP<EventMinutesService>
    {
        private readonly IServiceprovider _serviceprovider;
        private readonly BodyRoleService _roleSrv;
        private readonly EventAttendanceService _eventAttendanceService;

        public EventMinutesService(BodyRoleService roleSrv, ILogger<EventMinutesService> logger, M365AuthHelper auth, IServiceprovider serviceprovider, EventAttendanceService eventAttendanceService) : base(logger, auth)
        {
            _serviceprovider = serviceprovider;
            _roleSrv = roleSrv;
            _eventAttendanceService = eventAttendanceService;
        }

        public async Task<EventMinutesInformation> GetMinutesInformationByEvent(IPnPContext ctx, string bodyId, string sharedEventId, bool fromArchive = false)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var minutesInfoDTO = await GetDTO(bodyCtx, bodyId, sharedEventId, fromArchive);

                IEnumerable<EventUserAttendance> attendances = [];

                if (!fromArchive)
                {
                    attendances = await _eventAttendanceService.GetMinutesApprodvers(ctx, bodyId, sharedEventId);
                }

                return Map(minutesInfoDTO, attendances);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error  GetMinutesInformation by '{bodyId}' in eventId '{sharedEventId}'", ex);
            }
        }

        public async Task SetMinutesInformation(IPnPContext ctx, string bodyId, string sharedEventId, EventMinutesInformation eventMinutesInfo)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var minutesInfoDAL = new EventMinutesInformationDALprovider(eventDAL);

                //1. prodpiedades
                eventMinutesInfo.MinutesStatus = string.ECNTy;
                eventMinutesInfo.IdMinutes = string.ECNTy;

                //2. Get Information
                var minutesInfoDTO = await Map(bodyCtx, sharedEventId, eventMinutesInfo);
                await minutesInfoDAL.AddOrUpdateForEvent(bodyCtx, sharedEventId, minutesInfoDTO);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Set Certification by '{bodyId}' in sharedEventId '{sharedEventId}'", ex);
            }
        }

        public async Task SetMinutesDocument(IPnPContext ctx, string bodyId, string sharedEventId, string uIdDocument)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                var eventDAL = new InConstructionEventsDALprovider();
                var minutesDocDAL = new EventMinutesDocumentDALprovider(eventDAL);

                //1. Change ContentTypeGet
                await minutesDocDAL.SetContentTypeMinutesDocument(bodyCtx, uIdDocument);

                //2. Update Information Minutes
                await SetNewMinutesInMinutesInformation(bodyCtx, bodyId, sharedEventId, uIdDocument);

            }
            catch (Exception ex)
            {
                throw new Exception($"Error SetMinutesDocument by '{bodyId}' with uIdDocument:'{uIdDocument}'", ex);
            }
        }

        #region Métodos Privados

        private async Task RestarApprodvalprodccess(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            //1. Volver a Generar las Tareas de Aprodbación
            var eventService = _serviceprovider.GetRequiredService<EventsService>();
            await eventService.SendRequestApprodvalsMinutes(ctx, bodyId, sharedEventId);
        }

        private async Task<EventMinutesInformationDTO> GetDTO(IPnPContext bodyCtx, string bodyId, string sharedEventId, bool fromArchive = false)
        {
            var eventsprovider = new EventDALprovider(await _roleSrv.GetDefaultEventSourceByBodyRole(bodyCtx, bodyId, fromArchive));

            var minutesInfoDAL = new EventMinutesInformationDALprovider(eventsprovider);

            var minutesInfoDTO = await minutesInfoDAL.GetBySharedEventId(bodyCtx, sharedEventId);

            return minutesInfoDTO;
        }

        private async Task SetNewMinutesInMinutesInformation(IPnPContext ctx, string bodyId, string sharedEventId, string uIdMinutesDocument)
        {
            try
            {
                var eventDAL = new InConstructionEventsDALprovider();
                var minutesInfoDAL = new EventMinutesInformationDALprovider(eventDAL);

                //1. properties                                             
                var minutesInfoDTO = await GetDTO(ctx, bodyId, sharedEventId);

                if (minutesInfoDTO == null)
                {
                    // New
                    minutesInfoDTO = new EventMinutesInformationDTO
                    {
                        EstadoMinutes = DALConstants.TaxonomyValuesIds.RequestStatus.Pending,
                        IdMinutes = uIdMinutesDocument,
                        StartDateMinutes = DateTime.Now
                    };

                    await minutesInfoDAL.AddOrUpdateForEvent(ctx, sharedEventId, minutesInfoDTO);

                    //Se genera el prodceso de aprodbación al subir por primera vez un Minutes
                    await RestarApprodvalprodccess(ctx, bodyId, sharedEventId);
                }
                else
                {
                    //Update
                    minutesInfoDTO.EstadoMinutes = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;
                    minutesInfoDTO.IdMinutes = uIdMinutesDocument;
                    minutesInfoDTO.StartDateMinutes = DateTime.Now;
                    await minutesInfoDAL.AddOrUpdateForEvent(ctx, sharedEventId, minutesInfoDTO);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Set SetNewMinutes by '{bodyId}' in sharedEventId '{sharedEventId}'", ex);
            }
        }

        #endregion



        #region  Mapeos

        private EventMinutesInformation Map(EventMinutesInformationDTO item, IEnumerable<EventUserAttendance> attendances)
        {
            var result = new EventMinutesInformation
            {
                Attendances = attendances
            };

            if (item is null)
                return result;

            result.AgreementsInformation = item.InfAcuerdos;
            result.AttendanceInformation = item.InfAttendance;
            result.DocumentationInformation = item.InfDocumentacion;
            result.AgendaInformation = item.InfAgendaItem;
            result.MinutesStatus = item.EstadoMinutes;
            result.IdMinutes = item.IdMinutes;
            result.StartDateMinutes = item.StartDateMinutes;

            return result;
        }




        private async Task<EventMinutesInformationDTO> Map(IPnPContext ctx, String shareEventId, EventMinutesInformation item)
        {
            var eventDAL = new InConstructionEventsDALprovider();
            var minutesInfoDAL = new EventMinutesInformationDALprovider(eventDAL);

            var rDTO = await minutesInfoDAL.GetBySharedEventId(ctx, shareEventId);

            if (rDTO is null)
            {
                rDTO = new EventMinutesInformationDTO();
                rDTO.EstadoMinutes = DALConstants.TaxonomyValuesIds.RequestStatus.Pending;
            }


            rDTO.InfAcuerdos = item.AgreementsInformation;
            rDTO.InfAttendance = item.AttendanceInformation;
            rDTO.InfDocumentacion = item.DocumentationInformation;
            rDTO.InfAgendaItem = item.AgendaInformation;

            return rDTO;
        }

        #endregion

    }
}
