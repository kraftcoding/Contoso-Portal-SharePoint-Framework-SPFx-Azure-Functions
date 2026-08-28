using System.Xml.Linq;
using Microsoft.IdentityModel.Tokens;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAO.ConfigDepartments;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Model.profile;
using System.Reflection;
using PnP.Core.Model.SharePoint;
using Contoso.Portal.Data.DAO.DemoEntities;


namespace Contoso.Portal.Common
{

    public class InformationToXML
    {

        public static async Task<XDocument> RetrieveXMLAsync(string[]? files, TaxonomyTranslationService taxonomyService, Event? eventObj, EventAgreement? agreementObj, (string, string)[] votingInformation,
                                                            string urlDocuments, string urlMeeting)
        {
            string bodyId = eventObj?.BodyId ?? string.ECNTy;
            string eventId = eventObj?.Id ?? string.ECNTy;
            string agreementId = agreementObj?.Id.ToString() ?? string.ECNTy;

            List<string> guidsStr = [eventObj?.BodyTypeId!, eventObj?.BodyNameId!, eventObj?.AttendanceTypeId!, eventObj?.StatusId!, eventObj?.EventToolId!, agreementObj?.StatusId!];
            foreach (var voting in votingInformation)
            {
                guidsStr.Add(voting.Item1);
                guidsStr.Add(voting.Item2);
            }
            Guid[] guids = BuildGuidForTaxonomy(guidsStr);
            var taxonomyLabels = await taxonomyService.GetLabelsByIdAsync(guids, string.ECNTy);
            var status = agreementObj?.AgendaItemTypeId == DALConstants.TaxonomyValuesIds.AgendaItemType.Informative ? "Informativo" : C(GetValueTaxonomyById(taxonomyLabels, agreementObj?.StatusId!));

            try
            {
                XNamespace ns = "urn:Contoso/datasets/event";
                XDocument response = new XDocument(
                                            new XDeclaration("1.0", "utf-8", null),
                                            new XElement(ns + "data-set",
                                                new XElement("body",
                                                    new XElement("Id", C(bodyId)),
                                                    new XElement("Type", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyTypeId!))),
                                                    //new XElement("Type_Id", eventObj?.BodyTypeId),
                                                    new XElement("Name", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyNameId!))),
                                                    //new XElement("Name_Id", eventObj?.BodyTypeId),
                                                    new XElement("NowUniversalDateTime", C(DateTimeOperations.GetStringPrintFromDatetime(DateTime.UtcNow.ToUniversalTime()))),

                                                    new XElement("event",
                                                        new XElement("Id", C(eventId)),
                                                        new XElement("Title", C(eventObj?.Title)),
                                                        new XElement("Description", C(eventObj?.Description)),
                                                        new XElement("StartDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.StartDate))),
                                                        new XElement("EndDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.EndDate))),
                                                        new XElement("Location", C(eventObj?.Location)),
                                                        new XElement("LocationDetails", C(eventObj?.LocationDetails)),
                                                        new XElement("MeetingTooUrl", C(urlMeeting)),
                                                        new XElement("AttendanceType", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.AttendanceTypeId!))),
                                                        //new XElement("AttendanceType_Id", eventObj?.BodyTypeId),
                                                        new XElement("Status", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.StatusId!))),
                                                        //new XElement("Status_Id", eventObj?.StatusId),
                                                        new XElement("EventTool", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.EventToolId!))),
                                                        //new XElement("EventTool_Id", eventObj?.EventToolId),
                                                        new XElement("LastPublished", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.LastPublished))),

                                                        new XElement("agreement",
                                                            new XElement("Id", C(agreementId)),
                                                            new XElement("Title", C(agreementObj?.Title)),
                                                            new XElement("Description", C(agreementObj?.Description)),
                                                            new XElement("Status", string.IsNullOrECNTy(status) ? " " : status),
                                                            //new XElement("Status_Id", eventObj?.BodyTypeId),
                                                            new XElement("Order", C(agreementObj?.Order)),
                                                            new XElement("RelatedDocumentsUrl", C(urlDocuments)),
                                                            new XElement("Documents",
                                                                files?.Select(file => new XElement("Doc",
                                                                        new XElement("nameDoc", ReplaceName(file))))
                                                                    ),
                                                            new XElement("Votations", votingInformation?.Where(vot => !(vot.Item2 == null || vot.Item2 == string.ECNTy)).Select(votation => new XElement("Votation",
                                                                        new XElement("Retestsentation", C(GetValueTaxonomyById(taxonomyLabels, votation.Item1)) + ":"),
                                                                        new XElement("Vote", (votation.Item2 == null || votation.Item2 == string.ECNTy) ? "Pendiente" : C(GetValueTaxonomyById(taxonomyLabels, votation.Item2)))))
                                                                    ))

                                                    )
                                                )
                                            )
                                        );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }

        public static async Task<XDocument> RetrieveXMLAsync(TaxonomyTranslationService taxonomyService,
                                                            Event? eventObj, EventAgendaItem[]? agendasObj,
                                                            EventUserAttendance[]? attendancesObj, Dictionary<string, profileInfo> attendancesSharePointObj,
                                                            EventAgreement[]? agreementsObj, EventMinutesInformation? minutesObj,
                                                            string urlDocuments, string urlMeeting, IEnumerable<Data.DAO.Event.EventPublishedDocumentDAO> publishedDocuments, Dictionary<string, Summary> summaries)
        {
            // Retrieve all taxonomies for objects: event, agendas, event, agreements, minutes
            List<string> guidsStr = [eventObj?.BodyTypeId!, eventObj?.BodyNameId!, eventObj?.AttendanceTypeId!, eventObj?.StatusId!, eventObj?.EventToolId!, minutesObj?.MinutesStatus!];
            if (agendasObj != null)
            {
                foreach (var agenda in agendasObj)
                {
                    guidsStr.Add(agenda.AgendaItemType ?? string.ECNTy);
                }
            }
            if (agreementsObj != null)
            {
                foreach (var agreement in agreementsObj)
                {
                    guidsStr.Add(agreement.StatusId);
                }
            }
            if (attendancesObj != null)
            {
                foreach (var attendance in attendancesObj)
                {
                    guidsStr.Add(attendance.RequestStatusId);
                    guidsStr.Add(attendance.AttendanceTypeId);
                }
            }

            Guid[] guids = BuildGuidForTaxonomy(guidsStr);
            var allTaxonomyLabels = await taxonomyService.GetLabelsByIdAsync(guids, string.ECNTy);


            try
            {
                XNamespace ns = "urn:Contoso/datasets/event";
                XDocument response = new XDocument(
                                        new XDeclaration("1.0", "utf-8", null),
                                        new XElement(ns + "data-set",
                                            new XElement("body",
                                                new XElement("Id", C(eventObj?.BodyId)),
                                                new XElement("Type", C(GetValueTaxonomyById(allTaxonomyLabels, eventObj?.BodyTypeId!))),
                                                //new XElement("Type_Id", eventObj?.BodyTypeId),
                                                new XElement("Name", C(GetValueTaxonomyById(allTaxonomyLabels, eventObj?.BodyNameId!))),
                                                //new XElement("Name_Id", eventObj?.BodyTypeId),

                                                new XElement("event",
                                                    new XElement("Id", C(eventObj?.Id)),
                                                    new XElement("Title", C(eventObj?.Title)),
                                                    new XElement("Description", C(eventObj?.Description)),
                                                    new XElement("StartDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("BeginDate", C(DateTimeOperations.GetDateStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("BeginTime", C(DateTimeOperations.GetTimeStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("EndDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.EndDate))),
                                                    new XElement("Location", C(eventObj?.Location)),
                                                    new XElement("LocationDetails", C(eventObj?.LocationDetails)),
                                                    new XElement("MeetingTooUrl", C(urlMeeting)),
                                                    new XElement("AttendanceType", C(GetValueTaxonomyById(allTaxonomyLabels, eventObj?.AttendanceTypeId!))),
                                                    //new XElement("AttendanceType_Id", eventObj?.BodyTypeId),
                                                    new XElement("Status", C(GetValueTaxonomyById(allTaxonomyLabels, eventObj?.StatusId!))),
                                                    //new XElement("Status_Id", eventObj?.StatusId),
                                                    new XElement("EventTool", C(GetValueTaxonomyById(allTaxonomyLabels, eventObj?.EventToolId!))),
                                                    //new XElement("EventTool_Id", eventObj?.EventToolId),
                                                    new XElement("LastPublished", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.LastPublished))),

                                                    new XElement("RelatedDocumentsUrl", C(urlDocuments)),

                                                     new XElement("Documents",
                                                                publishedDocuments?.Select(file => new XElement("Doc",
                                                                        new XElement("nameDoc", ReplaceName(file.FileLeafRef))))
                                                                    ),

                                                    new XElement("agendas",
                                                        agendasObj?.Select(agenda => new XElement("agenda",
                                                                new XElement("Id", C(agenda.Id.ToString())),
                                                                new XElement("Title", " - " + C(agenda.Title)),
                                                                new XElement("Description", C(agenda.Description)),
                                                                new XElement("Duration", C(CultureOperations.DoubleAsIntToString(agenda.Duration))),
                                                                new XElement("Order", C(agenda.OrderFormatted)),
                                                                new XElement("AgendaItemType", C(GetValueTaxonomyById(allTaxonomyLabels, agenda.AgendaItemType!)))
                                                            )
                                                        )
                                                    ),

                                                    GetUsersAttendanceVote(allTaxonomyLabels, attendancesObj, attendancesSharePointObj),
                                                    GetUsersAttendanceGuest(allTaxonomyLabels, attendancesObj, attendancesSharePointObj),
                                                    GetUsersAttendanceVoteDelegate(allTaxonomyLabels, attendancesObj, attendancesSharePointObj),
                                                    GetUsersAttendanceAttendanceDelegate(allTaxonomyLabels, attendancesObj, attendancesSharePointObj),

                                                    new XElement("agreements",
                                                        agreementsObj?.Select(agreement => new XElement("agreement",
                                                                new XElement("Id", C(agreement.Id.ToString())),
                                                                new XElement("Title", " - " + C(agreement.Title)),
                                                                new XElement("Description", C(agreement.Description)),
                                                                new XElement("Status", C(GetValueTaxonomyById(allTaxonomyLabels, agreement.StatusId))),
                                                                new XElement("StatusTitle", "Estado del Acuerdo:"),
                                                                new XElement("Order", C(agreement.Order)),
                                                                new XElement("Summary",
                                                                    new XElement("Pending", summaries.TryGetValue(agreement.Id, out Summary? value) ? value.Pending.ToString() : string.ECNTy),
                                                                    new XElement("Approdve", summaries.TryGetValue(agreement.Id, out Summary? value1) ? value1.Approdve.ToString() : string.ECNTy),
                                                                    new XElement("Abstention", summaries.TryGetValue(agreement.Id, out Summary? value2) ? value2.Abstention.ToString() : string.ECNTy),
                                                                    new XElement("Reject", summaries.TryGetValue(agreement.Id, out Summary? value3) ? value3.Reject.ToString() : string.ECNTy)
                                                            )
                                                        )
                                                        )
                                                    ),

                                                    new XElement("minutes",
                                                        new XElement("IdMinutes", C(minutesObj?.IdMinutes)),
                                                        new XElement("AgendaInformation", C(minutesObj?.AgendaInformation)),
                                                        new XElement("AgreementsInformation", C(minutesObj?.AgreementsInformation)),
                                                        new XElement("AttendanceInformation", C(minutesObj?.AttendanceInformation)),
                                                        new XElement("DocumentationInformation", C(minutesObj?.DocumentationInformation)),
                                                        new XElement("MinutesStatus", C(GetValueTaxonomyById(allTaxonomyLabels, minutesObj?.MinutesStatus!)))
                                                    )

                                                )
                                            )
                                        )
                                    );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }

        public static async Task<XDocument> RetrieveXMLAsync(TaxonomyTranslationService taxonomyService, Event? eventObj,
                                                            EventUserAttendance[]? attendancesObj, Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            string eventId = eventObj?.Id ?? string.ECNTy;

            List<string> guidsStr = [eventObj?.BodyTypeId!, eventObj?.BodyNameId!, eventObj?.EventToolId!, eventObj?.AttendanceTypeId!, eventObj?.StatusId!];
            if (attendancesObj != null)
            {
                foreach (var attendance in attendancesObj)
                {
                    guidsStr.Add(attendance.RequestStatusId);
                    guidsStr.Add(attendance.AttendanceTypeId);
                }
            }
            Guid[] guids = BuildGuidForTaxonomy(guidsStr);
            var taxonomyLabels = await taxonomyService.GetLabelsByIdAsync(guids, string.ECNTy);

            try
            {
                XNamespace ns = "urn:Contoso/datasets/event";
                XDocument response = new XDocument(
                                        new XDeclaration("1.0", "utf-8", null),
                                        new XElement(ns + "data-set",
                                            new XElement("body",
                                                    new XElement("Id", eventObj?.BodyId),
                                                    new XElement("Type", GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyTypeId!)),
                                                    new XElement("Name", GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyNameId!)),
                                                    new XElement("NowUniversalDateTime", C(DateTimeOperations.GetStringPrintFromDatetime(DateTime.UtcNow.ToUniversalTime()))),

                                                new XElement("event",
                                                    new XElement("Id", C(eventObj?.Id)),
                                                    new XElement("Title", C(eventObj?.Title)),
                                                    new XElement("Description", C(eventObj?.Description)),
                                                    new XElement("StartDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.StartDate))),
                                                GetUsersAttendancetestsent(taxonomyLabels, attendancesObj, attendancesSharePointObj)
                                                ,
                                                GetUsersAttendanceAbsent(taxonomyLabels, attendancesObj, attendancesSharePointObj))
                                            )
                                        )
                                    );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }


        public static async Task<XDocument> RetrieveXMLAsync(TaxonomyTranslationService taxonomyService, Event? eventObj,
                                                    EventAgendaItem[]? agendasObj, string urlMeeting)
        {
            string eventId = eventObj?.Id ?? string.ECNTy;
            List<string> guidAgTypes = agendasObj.Select(ag => ag.AgendaItemType).ToHashSet().ToList();
            List<string> guidsStr = [eventObj?.BodyTypeId!, eventObj?.BodyNameId!, .. guidAgTypes];
            Guid[] guids = BuildGuidForTaxonomy(guidsStr);
            var taxonomyLabels = await taxonomyService.GetLabelsByIdAsync(guids, string.ECNTy);

            try
            {
                XNamespace ns = "urn:Contoso/datasets/event";
                XDocument response = new XDocument(
                                        new XDeclaration("1.0", "utf-8", null),
                                        new XElement(ns + "data-set",
                                            new XElement("body",
                                                new XElement("Id", C(eventObj?.BodyId)),
                                                new XElement("Type", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyTypeId!))),
                                                //new XElement("Type_Id", eventObj?.BodyTypeId),
                                                new XElement("Name", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.BodyNameId!))),
                                                //new XElement("Name_Id", eventObj?.BodyTypeId),

                                                new XElement("event",
                                                    new XElement("Id", C(eventObj?.Id)),
                                                    new XElement("Title", C(eventObj?.Title)),
                                                    new XElement("Description", C(eventObj?.Description)),
                                                    new XElement("StartDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("BeginDate", C(DateTimeOperations.GetDateStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("BeginTime", C(DateTimeOperations.GetTimeStringPrintFromDatetime(eventObj?.StartDate))),
                                                    new XElement("EndDate", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.EndDate))),
                                                    new XElement("Location", C(eventObj?.Location)),
                                                    new XElement("LocationDetails", C(eventObj?.LocationDetails)),
                                                    new XElement("MeetingTooUrl", C(urlMeeting)),
                                                    new XElement("AttendanceType", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.AttendanceTypeId!))),
                                                    //new XElement("AttendanceType_Id", eventObj?.BodyTypeId),
                                                    new XElement("Status", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.StatusId!))),
                                                    //new XElement("Status_Id", eventObj?.StatusId),
                                                    new XElement("EventTool", C(GetValueTaxonomyById(taxonomyLabels, eventObj?.EventToolId!))),
                                                    //new XElement("EventTool_Id", eventObj?.EventToolId),
                                                    new XElement("LastPublished", C(DateTimeOperations.GetStringPrintFromDatetime(eventObj?.LastPublished))),

                                                    //new XElement("RelatedDocumentsUrl", C(urlDocuments)),

                                                    new XElement("agendas",
                                                        agendasObj?.Select(agenda => new XElement("agenda",
                                                                new XElement("Id", C(agenda.Id.ToString())),
                                                                new XElement("Title", "- " + C(agenda.Title)),
                                                                new XElement("Description", C(agenda.Description)),
                                                                new XElement("Duration", C(CultureOperations.DoubleAsIntToString(agenda.Duration))),
                                                                new XElement("Order", C(agenda.OrderFormatted)),
                                                                new XElement("AgendaItemType", C(GetValueTaxonomyById(taxonomyLabels, agenda.AgendaItemType!)))

                                                    )
                                                )
                                            )
                                        )
                                    )
                                )
                            );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }


        public static XDocument RetrieveXML(ConfigDepartmentsDAO bodyInformation)
        {
            try
            {
                XNamespace ns = "urn:Contoso/datasets/body";
                XDocument response = new XDocument(
                                        new XDeclaration("1.0", "utf-8", null),
                                        new XElement(ns + "data-set",
                                            new XElement("body",
                                                new XElement("Id", C(bodyInformation.Title)),
                                                new XElement("dir", C(bodyInformation.dir)),
                                                new XElement("SIA", C(bodyInformation.SIA)),
                                                new XElement("Materia", C(bodyInformation.Materia)),
                                                new XElement("Abreviatura", C(bodyInformation.Abreviatura)),
                                                new XElement("Activo", C(bodyInformation.Activo.ToString())),
                                                new XElement("FechaConstitucion", C(bodyInformation.FechaConstitucion.ToString())),
                                                new XElement("FechaExtincion", C(bodyInformation.FechaExtincion.ToString())),
                                                new XElement("Observaciones", C(bodyInformation.Observaciones)),
                                                new XElement("IdEmailBodies", C(bodyInformation.IdEmailBodies.ToString())),
                                                new XElement("IdActBodies", C(bodyInformation.IdActBodies.ToString())),
                                                new XElement("IdCertificateBodies", C(bodyInformation.IdCertificateBodies.ToString())),
                                                //new XElement("IdCertificateBody", C(bodyInformation.IdCertificateBody.ToString())),
                                                new XElement("IdAgendaBodies", C(bodyInformation.IdAgendaBodies.ToString())),
                                                new XElement("IdAttendanceBodies", C(bodyInformation.IdAttendanceBodies.ToString())),
                                                new XElement("DiasAprodbacionMinutes", C(bodyInformation.DiasAprodbacionMinutes.ToString())),
                                                new XElement("DocumentSetDescription", C(bodyInformation.DocumentSetDescription)),
                                                new XElement("Intersectorial", C(bodyInformation.Intersectorial.ToString())),
                                                new XElement("IdentificadorConferencia", C(bodyInformation.IdentificadorConferencia)),
                                                new XElement("Department", C(bodyInformation.Department?.Label)),
                                                new XElement("TipoDepartment", C(bodyInformation.TipoDepartment?.Label)),
                                                new XElement("Secretaria", C(bodyInformation.Secretaria?.Label)),
                                                new XElement("Division", C(bodyInformation.Division?.Label)),
                                                new XElement("Period", C(bodyInformation.Period?.Label)),
                                                new XElement("BusinessArea", C(bodyInformation.BusinessArea?.Label)),
                                                new XElement("TipoMembresia", C(bodyInformation.TipoMembresia?.Label)),
                                                new XElement("FechaActual", C(CultureOperations.DateTimeToLargeDateString(CultureOperations.GetCurrentSpanishTime(), DALConstants."en-US")))
                                            )
                                        )
                                    );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }

        public static XDocument RetrieveXML(DemoEntityDAO bodyInformation)
        {
            try
            {
                XNamespace ns = "urn:Contoso/datasets/DemoEntity";
                XDocument response = new XDocument(
                    new XDeclaration("1.0", "utf-8", null),
                    new XElement(ns + "data-set",
                        new XElement("body",
                            // Identificadores base (si hereda de BaseSPListItemDAO)
                            new XElement("Id", C(bodyInformation.ID.ToString())),
                            new XElement("Denominacion", C(bodyInformation.Title)),

                            // Campos específicos _EMD
                            new XElement("Folder", C(bodyInformation.Folder_EMD)),
                            new XElement("SectionId", C(bodyInformation.Section_EMD?.LookupValue)),
                            new XElement("ScopeTerritorialId", C(bodyInformation.ScopeTerritorial_EMD?.LookupValue)),

                            new XElement("FechaConstitucion", C(bodyInformation.FechaConstitucion_EMD.HasValue ?
                                CultureOperations.DateTimeToLargeDateString(bodyInformation.FechaConstitucion_EMD.Value, DALConstants."en-US") : string.ECNTy)),

                            new XElement("FechaInscripcion", C(bodyInformation.FechaInscripcion_EMD.HasValue ?
                                CultureOperations.DateTimeToLargeDateString(bodyInformation.FechaInscripcion_EMD.Value, DALConstants."en-US") : string.ECNTy)),

                            new XElement("DatosInscripcion", C(bodyInformation.DatosInscripcion_EMD)),
                            new XElement("DomicilioSocial", C(bodyInformation.DomicilioSocial_EMD)),
                            new XElement("NIF", C(bodyInformation.NIF_EMD)),
                            new XElement("Objeto", C(bodyInformation.Objeto_EMD)),
                            new XElement("EntidadesIntegrantes", C(bodyInformation.EntidadesIntegrantes_EMD)),
                            new XElement("IdentidadTitulares", C(bodyInformation.IdentidadTitulares_EMD)),

                            // Datos del Retestsentante
                            new XElement("RetestsentanteNombre", C(bodyInformation.RetestsentanteNombre_EMD)),
                            new XElement("RetestsentantePrimerApellido", C(bodyInformation.RetestsentantePrimerApellido_EMD)),
                            new XElement("RetestsentanteSegundoApellido", C(bodyInformation.RetestsentanteSegundoApellido_EMD)),
                            new XElement("RetestsentanteFechaNacimiento", C(bodyInformation.RetestsentanteFechaNacimiento_EMD.HasValue ?
                                CultureOperations.DateTimeToLargeDateString(bodyInformation.RetestsentanteFechaNacimiento_EMD.Value, DALConstants."en-US") : string.ECNTy)),
                            new XElement("RetestsentanteNacionalidad", C(bodyInformation.RetestsentanteNacionalidad_EMD)),
                            new XElement("RetestsentanteTipoDocument", C(bodyInformation.RetestsentanteTipoDocument_EMD)),
                            new XElement("RetestsentanteID", C(bodyInformation.RetestsentanteID_EMD)),
                            new XElement("RetestsentanteCargo", C(bodyInformation.RetestsentanteCargo_EMD)),

                            // Datos del Solicitante y Expediente
                            new XElement("NumExpedienteACCEDA", C(bodyInformation.NumExpedienteACCEDA_EMD)),
                            new XElement("SolicitanteNombre", C(bodyInformation.SolicitanteNombre_EMD)),
                            new XElement("SolicitanteApellido", C(bodyInformation.SolicitanteApellido_EMD)),
                            new XElement("SolicitanteSegundoApellido", C(bodyInformation.SolicitanteSegundoApellido_EMD)),
                            new XElement("SolicitanteTipoDocument", C(bodyInformation.SolicitanteTipoDocument_EMD)),
                            new XElement("SolicitanteNumIdentificacion", C(bodyInformation.SolicitanteNumIdentificacion_EMD)),

                            // Contacto y Estado
                            new XElement("ContactoEmail", C(bodyInformation.ContactoEmail_EMD)),
                            new XElement("ContactoTelefono", C(bodyInformation.ContactoTelefono_EMD)),
                            new XElement("CancellationReason", C(bodyInformation.CancellationReason_EMD?.LookupValue)),
                            new XElement("NumExpedienteCancelacion", C(bodyInformation.NumExpedienteCancelacion_EMD)),
                            new XElement("Estado", C(bodyInformation.Estado_EMD)),
                            new XElement("Asiento", C(bodyInformation.Asiento_EMD)),

                            // Fecha de generación
                            new XElement("FechaActual", C(CultureOperations.DateTimeToLargeDateString(CultureOperations.GetCurrentSpanishTime(), DALConstants."en-US")))
                        )
                    )
                );
                return response;
            }
            catch (Exception ex)
            {
                // He mantenido el nombre del método en el error según tu snippet
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXML)}", ex);
            }
        }


        public static XDocument RetrieveXML(ConfigEXTERNALBodiesDAO bodyInformation)
        {
            try
            {
                XNamespace ns = "urn:Contoso/datasets/body";
                XDocument response = new XDocument(
                                        new XDeclaration("1.0", "utf-8", null),
                                        new XElement(ns + "data-set",
                                            new XElement("body",
                                                new XElement("Id", C(bodyInformation.Title)),
                                                new XElement("CodigoDepartment", C(bodyInformation.CodigoDepartment?.ToUpper())),
                                                new XElement("Observaciones", C(bodyInformation.Observaciones?.ToUpper())),
                                                new XElement("Abreviatura", C(bodyInformation.Abreviatura?.ToUpper())),
                                                new XElement("SecretariaText", C(bodyInformation.SecretariaText?.ToUpper())),
                                                new XElement("DocumentSetDescription", C(bodyInformation.DocumentSetDescription?.ToUpper())),
                                                new XElement("dir", C(bodyInformation.dir?.ToUpper())),
                                                new XElement("SIA", C(bodyInformation.SIA?.ToUpper())),
                                                new XElement("Department", C(bodyInformation.Department?.Label?.ToUpper())),
                                                new XElement("TipoDepartmentEXTERNAL", C(bodyInformation.TipoDepartmentLookup?.LookupValue?.ToUpper())),
                                                new XElement("BusinessArea", C(bodyInformation.BusinessArea?.Label?.ToUpper())),
                                                new XElement("StatusDepartment", C(bodyInformation.StatusLookup?.LookupValue?.ToUpper())),
                                                new XElement("Division", C(bodyInformation.DivisionLookup?.LookupValue?.ToUpper())),
                                                new XElement("InscritoDepartment", C(bodyInformation.InscritoDepartment.ToString())),
                                                new XElement("DepartmentAdscripcion", C(bodyInformation.DepartmentAdscripcionTax?.Label ?? string.ECNTy)),
                                                new XElement("FechaCreacionDepartment", C(bodyInformation.FechaCreacionDepartment.HasValue ? bodyInformation.FechaCreacionDepartment.ToString() : string.ECNTy)),
                                                new XElement("Activo", C(bodyInformation.Activo.ToString())),
                                                new XElement("Intersectorial", C(bodyInformation.Intersectorial.ToString())),
                                                new XElement("FechaConstitucion", C(bodyInformation.FechaConstitucion != null ? ((DateTime)bodyInformation.FechaConstitucion).ToString("dd/MM/yyyy") : "")),
                                                new XElement("FechaExtincion", C(bodyInformation.FechaExtincion != null ? ((DateTime)bodyInformation.FechaExtincion).ToString("dd/MM/yyyy") : "")),
                                                new XElement("FechaInscripcion", C(bodyInformation.FechaInscripcion != null ? ((DateTime)bodyInformation.FechaInscripcion).ToString("dd/MM/yyyy") : "")),
                                                new XElement("FechaActual", C(CultureOperations.DateTimeToLargeDateString(CultureOperations.GetCurrentSpanishTime(), DALConstants."en-US")))
                                            )
                                        )
                                    );

                return response;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error {nameof(InformationToXML) + "." + nameof(RetrieveXMLAsync)}", ex);
            }
        }

        #region Private method
        private static readonly string _SECNTY_ = " ";
        private enum UsersMode { MINUTES, MINUTES_GEST, ATTENDANCE_LIST };
        private enum DelegationMode { ATTENDANCE, VOTE };

        private static string C(string? value) => string.IsNullOrWhiteSpace(value) ? _SECNTY_ : value;

        private static Guid[] BuildGuidForTaxonomy(List<string> guidsStr)
        {
            return guidsStr.Where(s => !string.IsNullOrWhiteSpace(s) && Guid.TryParse(s, out _))
                            .Distinct()
                            .Select(s => new Guid(s))
                            .ToArray();
        }

        private static string GetValueTaxonomyById(Dictionary<Guid, string> taxonomies, string guidStr)
        {
            Guid searchGuid;
            if (!Guid.TryParse(guidStr, out searchGuid))
            {
                return _SECNTY_;
            }
            return taxonomies.GetValueOrDefault(searchGuid) ?? _SECNTY_;
        }

        private static string ReplaceName(string file)
        {
            int lastPointReplace = file.LastIndexOf('.');
            string nameDoc;
            if (lastPointReplace != -1 && lastPointReplace != file.Length - 1)
            {
                nameDoc = file.Substring(0, lastPointReplace);
            }
            else
            {
                nameDoc = file;
            }
            return nameDoc;
        }

        private static XElement GetUsersAttendancetestsent(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesMapToXElement(taxonomies, "attendancestestsent",
                                            attendances?.Where(usr => usr.HasAttended == true).ToArray(),
                                            attendancesSharePointObj, UsersMode.MINUTES);
        }

        private static XElement GetUsersAttendanceAbsent(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesMapToXElement(taxonomies, "attendancesAbsent",
                                            attendances?.Where(usr => usr.HasAttended == false).ToArray(),
                                            attendancesSharePointObj, UsersMode.MINUTES);
        }


        private static XElement GetUsersAttendanceVote(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesMapToXElement(taxonomies, "attendancesVote",
                                            attendances?.Where(usr => usr.IsGuest == false).ToArray(),
                                            attendancesSharePointObj, UsersMode.MINUTES);
        }
        private static XElement GetUsersAttendanceGuest(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesMapToXElement(taxonomies, "attendanceGuest",
                                            attendances?.Where(usr => usr.IsGuest == true).ToArray(),
                                            attendancesSharePointObj, UsersMode.MINUTES_GEST);
        }

        private static XElement GetUsersAttendanceVoteDelegate(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesDelegateMapToXElement("attendancesVoteDelegate", taxonomies,
                                            attendances?.Where(usr => !string.IsNullOrWhiteSpace(usr.DelegationUserVotePrincipalName)).ToArray(),
                                            attendancesSharePointObj, DelegationMode.VOTE);
        }
        private static XElement GetUsersAttendanceAttendanceDelegate(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            return AtendancesDelegateMapToXElement("attendancesGoDelegate", taxonomies,
                                            attendances?.Where(usr => !string.IsNullOrWhiteSpace(usr.DelegationUserVotePrincipalName)).ToArray(),
                                            attendancesSharePointObj, DelegationMode.ATTENDANCE);
        }

        private static XElement GetUsersAttendance(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                    Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            // Single unseparated listing
            return AtendancesMapToXElement(taxonomies, "attendances", attendances, attendancesSharePointObj, UsersMode.ATTENDANCE_LIST);
        }

        private static XElement GetUsersAttendanceConvo(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                  Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            attendances = attendances?.Where(att => att.RequestStatusId != DALConstants.TaxonomyValuesIds.RequestStatus.Accepted).ToArray();
            return AtendancesMapToXElement(taxonomies, "attendancesConvo", attendances, attendancesSharePointObj, UsersMode.ATTENDANCE_LIST);
        }

        private static XElement GetUsersAttendanceAccept(Dictionary<Guid, string> taxonomies, EventUserAttendance[]? attendances,
                                                  Dictionary<string, profileInfo> attendancesSharePointObj)
        {
            attendances = attendances?.Where(att => att.RequestStatusId == DALConstants.TaxonomyValuesIds.RequestStatus.Accepted).ToArray();
            return AtendancesMapToXElement(taxonomies, "attendancesAccept", attendances, attendancesSharePointObj, UsersMode.ATTENDANCE_LIST);
        }


        private static XElement AtendancesMapToXElement(Dictionary<Guid, string> taxonomies,
                                                        string nameParentNode,
                                                        EventUserAttendance[]? attendances,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj,
                                                        UsersMode mode)
        {
            return new XElement(nameParentNode,
                attendances?.Select(attendance => AtendanceMapToXElement(taxonomies, attendance, attendancesSharePointObj, mode))
            );
        }
        private static XElement AtendanceMapToXElement(Dictionary<Guid, string> taxonomies, EventUserAttendance attendance,
                                                        Dictionary<string, profileInfo> attendancesSharePointObj, UsersMode mode)
        {
            profileInfo? infoUsr;
            string infoFirstName, infoLastName, infoEmail, infoCellPhone, infoJobTitle, infoOffice;
            if (attendancesSharePointObj.TryGetValue(attendance.UserPrincipalName, out infoUsr))
            {
                infoFirstName = infoUsr.FirstName;
                infoLastName = infoUsr.LastName;
                infoEmail = infoUsr.Email;
                infoCellPhone = infoUsr.CellPhone;
                infoJobTitle = infoUsr.JobTitle;
                infoOffice = infoUsr.Office;
            }
            else
            {
                infoFirstName = infoLastName = infoEmail = infoCellPhone = infoJobTitle = infoOffice = string.ECNTy;
            }

            string voteDeleg = string.ECNTy, vote = string.ECNTy, typeUsr = string.ECNTy;
            switch (mode)
            {
                case UsersMode.MINUTES:
                    voteDeleg = "Delega voto: ";
                    break;
                case UsersMode.MINUTES_GEST:
                    break;
                case UsersMode.ATTENDANCE_LIST:
                    vote = "Vota: ";
                    voteDeleg = "Delega voto: ";
                    typeUsr = "Convocado como invitado: ";
                    break;
            }

            return new XElement("attendance",
                        // Information user Back
                        new XElement("AttendanceId", C(attendance.AttendanceId)),
                        new XElement("UserPrincipalName", C(attendance.UserPrincipalName)),
                        new XElement("Status", C(GetValueTaxonomyById(taxonomies, attendance.RequestStatusId))),
                        new XElement("RequestUserMessage", C(attendance.RequestUserMessage)),
                        new XElement("RequestStatusUpdate", C(DateTimeOperations.GetStringPrintFromDatetime(attendance.RequestStatusUpdate))),
                        new XElement("HasAttended", attendance.HasAttended != true ? " (no ha asistido)" : " "),
                        new XElement("AttendanceType", string.IsNullOrECNTy(attendance.AttendanceTypeId) ? "" : " - " + C(GetValueTaxonomyById(taxonomies, attendance.AttendanceTypeId))),
                        new XElement("Vote", attendance.Vote ? vote + "Si" : vote + "No"),
                        new XElement("IsGuest", attendance.IsGuest == true ? typeUsr + "Si" : typeUsr + "No"),
                        new XElement("VoteDelegation", attendance.VoteDelegation == null ? _SECNTY_ : (attendance.VoteDelegation == true ? voteDeleg + "Si" : voteDeleg + "No")),
                        new XElement("DelegationUserPrincipalName", C(attendance.DelegationUserPrincipalName)),
                        new XElement("DelegationUserVotePrincipalName", C(attendance.DelegationUserVotePrincipalName)),
                        // IsGuest
                        // AttendanceParentId
                        // Information user SharePoint
                        new XElement("FirstName", C(infoFirstName)),
                        new XElement("LastName", " " + C(infoLastName)),
                        new XElement("Email", C(infoEmail)),
                        new XElement("CellPhone", C(infoCellPhone)),
                        new XElement("JobTitle", C(infoJobTitle)),
                        new XElement("Office", C(infoOffice))
                    );
        }


        private static XElement AtendancesDelegateMapToXElement(string nodeName,
                                                                Dictionary<Guid, string> taxonomies,
                                                                EventUserAttendance[]? attendances,
                                                                Dictionary<string, profileInfo> attendancesSharePointObj,
                                                                DelegationMode delegationMode
                                                                )
        {
            return new XElement(nodeName,
                                    attendances?.Select(attendance => AtendanceDelegationMapToXElement(taxonomies, attendance, attendancesSharePointObj, delegationMode))
                                );
        }

        private static XElement AtendanceDelegationMapToXElement(Dictionary<Guid, string> taxonomies, EventUserAttendance attendance,
                                                                Dictionary<string, profileInfo> attendancesSharePointObj, DelegationMode delegationMode)
        {
            profileInfo? sourceUsr, destinationUsr;
            _ = attendancesSharePointObj.TryGetValue(attendance.UserPrincipalName, out sourceUsr);
            _ = attendancesSharePointObj.TryGetValue(attendance.DelegationUserPrincipalName, out destinationUsr);

            string action = string.ECNTy;
            switch (delegationMode)
            {
                case DelegationMode.VOTE: action = "ha delegado el voto en"; break;
                case DelegationMode.ATTENDANCE: action = "ha delegado la Attendance en"; break;
            }

            return new XElement("delegate",
                        new XElement("source", C(sourceUsr?.FirstName + " " + sourceUsr?.LastName)),
                        new XElement("info", action),
                        new XElement("destination", C(destinationUsr?.FirstName + " " + destinationUsr?.LastName))
                    );
        }

        #endregion
    }
} // namespace
