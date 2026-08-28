namespace Contoso.Portal.Data.DAL
{
    public static class DALConstants
    {
        public const string SPAppLoginName = "i:0i.t|00000003-0000-0ff1-ce00-000000000000|app@sharepoint"; // from https://pnp.github.io/pnpcore/using-the-sdk/taxonomy-intro.html#working-the-term-store
        public const string HiddenFolderName = "_hidden";

        public const string EventGroupprefix = "Attendees of ";

        public static class PnPContexts
        {
            public const string RootSiteAsSystem = "Default";
        }

        public static class ReplacedTokens
        {
            public static class BodyArticles
            {
                public const string MasculineArticle = "el";
                public const string FeminineArticle = "e la";
                public const string NeutralArticle = "e";
                public static readonly List<string> MasculineBodyTypeNouns = ["consejo", "grupo"];
                public static readonly List<string> FeminineBodyTypeNouns = ["conferencia", "comision"];
            }
            public const string DepartmentName = "{DepartmentName}";
            public const string MeetingTitle = "{MeetingTitle}";
            public const string AgendaTitle = "{AgendaTitle}";
            public const string PersonActionAuthor = "{PersonActionAuthor}";
            public const string PersonApprodver = "{PersonApprodver}";
            public const string UrgentNotificationBody = "{UrgentNotificationBody}";
            public const string StarDate = "{startDate}";
            public const string EndDate = "{endDate}";
            public const string PublishedVersionDocumentSetDescription = "{PublishedVersion.DocumentSetDescription}";
            public const string DetailURL = "{eventDetailsUrl}";
            public const string Location = "{location}";
            public const string RejectionReason = "{RejectionReason}";
            public const string DelegatedBy = "{DelegatedBy}";
            public const string BodyArticle = "{bodyArticle}";
            public const string AttendanceType = "{attendanceType}";
        }

        public static class Cache
        {
            public const string ConfigDepartments = "ConfigDepartments";

            public const string ConfigEXTERNALBodies = "ConfigDepartmentsR301cO";

            public const string NotificationMessages = "NOTIFICATION_MESSAGES";
        }

        public static class EmailBodyTypes
        {
            public const string Urgent = "UrgentCommunications";

            public const string ReservadaAgenda = "Reserved";
            public const string Publicada = "Published";

            public const string EnCelebracion = "InCelebration";

            public const string Finalizada = "Finished";

            public const string Reminder1H = "Reminder1H";

            public const string Reminder24H = "Reminder24H";

            public const string Archivada = "Archived";

            public const string Cancelled = "Cancelled";

            public const string InformacionDelegationAttendanceYVoto = "InformacionDelegationAttendanceYVoto";

            public const string InformacionDelegationAttendance = "InformacionDelegationAttendance";

            public const string InformacionDelegationVoto = "InformacionDelegationVoto";

            public const string RechazarDelegation = "RechazarDelegation";
        }



        public static class LocaleContoso
        {
            public const string esES = "es-ES";
            public const string euES = "eu-ES";
            public const string caES = "ca-ES";
            public const string glES = "gl-ES";

        }

        public static class NotificationsMessages
        {
            public static class Reminders
            {
                public const string OneHour = "1 h";
                public const string TwentyFourHours = "24 h";

            }
            public static class Tipo
            {
                public const string NoDefinida = "NoDefinida";
                public const string Error = "Error";
                public const string Sistema = "Sistema";
                public const string Meetings = "Meetings";
                public const string Alert = "Alert";
                public const string Request = "Request";
            }
            public static class Prioridad
            {
                public const string Low = "Low";
                public const string Half = "Half";
                public const string High = "High";
                public const string Urgent = "Urgent";
            }



            public static class Alert
            {
                public const string Urgent = "AlertsUrgent";
                public const string ErrorGenerico = "AlertsErrorGenerico";
                public const string AceptarDelegation = "AlertsAceptarDelegation";
                public const string InformacionDelegationAttendanceYVoto = "InformacionDelegationAttendanceYVoto";
                public const string InformacionDelegationAttendance = "InformacionDelegationAttendance";
                public const string InformacionDelegationVoto = "InformacionDelegationVoto";
                public const string RechazarDelegation = "RechazarDelegation";
                public const string MinutesValidada = "AlertsMinutesValidada";
                public const string RechazarMinutes = "AlertsRechazarMinutes";
                public const string PublicarMinutes = "AlertsPublicarMinutes";
                public const string AceptarModificacionMinutes = "AlertsAceptarModificacionMinutes";
                public const string RechazarModificacionMinutes = "AlertsRechazarModificacionMinutes";
                public const string NuevoCertificado = "AlertsNuevoCertificado";

            }

            public static class Meeting
            {
                public const string Nueva = "MeetingNueva";
                public const string Guardar = "MeetingGuardar";
                public const string Modificar = "MeetingModificar";
                public const string ModificarOrdenDia = "MeetingModificarOrdenDia";
                public const string ModificarCertificados = "MeetingModificarCertificados";
                public const string ModificarMinutes = "MeetingModificarMinutes";
                public const string Publicar = "MeetingPublicar";
                public const string EnCelebracion = "MeetingEnCelebracion";
                public const string Celebrada = "MeetingCelebrada";
                public const string Finalizada = "MeetingFinalizada";
                public const string Archivada = "MeetingArchivada";
                public const string Cancelada = "MeetingCancelada";
                public const string Reminder1H = "Reminder1H";
                public const string Reminder24H = "Reminder24H";
            }

            public static class Request
            {
                public const string Attendance = "RequestAttendance";
                public const string DelegarAttendance = "RequestDelegarAttendance";
                public const string DelegateVote = "RequestDelegateVote";
                public const string GeneracionCertificado = "RequestGeneracionCertificado";
                public const string ModificacionMinutes = "RequestModificacionMinutes";
            }
        }

        public static class SPLoginNameFormat
        {
            //https://pnp.github.io/pnpcore/using-the-sdk/security-users.html#sharepoint-users
            public const string ADGroup = "c:0t.c|tenant|";
            public const string M365Group = "c:0o.c|federateddirectoryclaimprovider|";
            public const string User = "i:0#.f|membership|";
        }

        public static class Userproperties
        {
            public static class Graph
            {
                public const string testferredLanguage = "testferredLanguage ";
                public const string OtherMails = "OtherMails";
                public const string Mail = "Mail";
                public const string JobTitle = "JobTitle";
                public const string Surname = "surname";
                public const string GivenName = "givenName";
                public const string OfficeLocation = "officeLocation";
                public const string BusinessPhones = "businessPhones";
                public const string DisplayName = "displayName";
                public const string AssignedLicenses = "assignedLicenses";
                public const string UserPrincipalName = "userPrincipalName";
                public const string UsageLocation = "usageLocation";
                public const string CompanyName = "companyName";
                public const string Department = "department";
                public const string EmployeeType = "employeeType";
                public const string MobilePhone = "mobilePhone";


            }

            public static class LicensesTypes
            {
                public const string E3 = "05e9a617-0261-4cee-bb44-138d3ef5d965";
                public const string E5 = "06ebc4ee-1bb5-47dd-8120-11324bc54e06";
                public const string E5Developer = "c42b9cae-ea4f-4ab7-9717-81576235ccac";
                public const string PowerAutomateFree = "f30db892-07e9-47e9-837c-80727f46fd3d";
                public const string PowerBI = "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235";
                public const string Teamstestmium = "36a0f3b3-adb5-49ea-bf66-762134cf063a";
                public const string PowerAppsDeveloper = "5b631642-bd26-49fe-bd20-1daaa972ef80";

            }

            public static class ParticipantsInfo
            {

                public const string Mail = "mail";

                public const string UPN = "userPrincipalName";

                public const string Identities = "identities";

                public static class Options
                {
                    public const string Federated = "federated";

                    public const string Internal = "userPrincipalName";

                    public const string ExternalAzure = "ExternalAzureAD";

                    public const string ExternalOther = "MicrosoftAccount";

                }

            }
        }

        public static class ContentTypeIds
        {
            public const string EventBase = "0x0120D520000639E2BEA58C8844A67C464647ED1630";
            public const string InConstructionEvent = "0x0120D520000639E2BEA58C8844A67C464647ED163001";
            public const string PublishedEvent = "0x0120D520000639E2BEA58C8844A67C464647ED163002";
            public const string ArchivedEvent = "0x0120D520000639E2BEA58C8844A67C464647ED163003";
            public const string AgendaItem = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020102";
            public const string MinutesInformation = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020105";
            public const string MinutesDocument = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020201";
            public const string ConfiguracionDepartments = "0x0120D520000639E2BEA58C8844A67C464647ED1640";
            public const string ConfiguracionDepartmentsEXTERNAL = "0x0120D52000A44E5324B75E844F854109CDB71AD8EE";
            public const string Folder = "0x0120";
            public const string NotificationConfig = "0x010100C2479835CA85CC4C899DC8456174BB55";
            public const string NotificationMessages = "0x010100C2479835CA85CC4C899DC8456174BB5501";
            public const string Notification = "0x010100F7FFABE611908A4EB5904C7F6AE5B435";
            public const string Acuerdo = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020101";
            public const string AgendaItem = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020102";
            public const string DocumentEnConstruccion = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC010202";
            public const string DocumentCertificacion = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020203";
            public const string DocumentDefinitivoPTCAAPP = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC010203";
            public const string DocumentEniPTCAAPP = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC010204";
            public const string VotationDetail = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010306";
            public const string ManagementRequest = "0x0100574224502F524E4AB6883DAACD3DC51D";
            public const string BusinessUnit = "0x010097E010D722B0454AB39A72CE3C836F7A";
            public const string processReport = "0x0101002C7CF9DC5C02B04F8193C84C6BFC5C39";

            public const string DemoEntity = "0x0120D5200047412C5763C0F44498C8B01E8FB1A6AE";
            public const string DocumentDemoEntity = "0x0101002E6D9650F2396148BB9B0DDF74F71A84";


            public static class Task
            {
                public const string Request = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC01020103";
                public const string RequestAttendance = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010301";
                public const string RequestDelegation = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010302";
                public const string RequestCertificacion = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010303";
                public const string RequestModificacion = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010304";
                public const string RequestAprodbacionMinutes = "0x0101006B46B0FAC7593A4E8EFEA5A5399F44AC0102010305";
            }
        }

        public static class ListsSiteRelativeUrls
        {
            public const string NotificationsToRead = "/Notifications";
            public const string NotificationsReaded = "/NotificationsHidden";
            public const string NotificationConfig = "/NotificationUserConfig";
            public const string InConstructionEvents = "/MeetingsConstruccion";
            public const string PublishedEvents = "/MeetingsFinales";
            public const string ArchivedEvents = "/MeetingsArchivadas";
            public const string StagingDocuments = "/DocumentsTemporales";
            public const string ManagementRequests = "/Lists/RequestesAdministracion";
            public const string BusinessUnit = "/Lists/BusinessAreaDivision";
            public const string processReports = "/processReports";

            public static class Home
            {
                public const string ConfigDepartments = "/ConfiguracionDepartments";
                public const string ConfigEXTERNALBodies = "/DepartmentsEXTERNAL";
            }

            public static class RegistroDemoEntities
            {
                public const string Entidades = "/DemoEntities";
                public const string Sections = "/Lists/SectionsDemoEntities";
                public const string Scopes = "/Lists/ScopesDemoEntities";
                public const string CancellationReasons = "/Lists/CancellationReasonsDemoEntities";
            }
        }

        public static class Fields
        {
            // Default fields
            public const string Id = "ID";
            public const string UniqueId = "UniqueId";
            public const string Title = "Title";
            public const string Created = "Created";
            public const string Modified = "Modified";
            public const string Author = "Author";
            public const string Editor = "Editor";
            public const string ContentTypeId = "ContentTypeId";
            public const string FileRef = "FileRef";
            public const string FileDirRef = "FileDirRef";
            public const string FileLeafRef = "FileLeafRef";
            public const string FileSize = "File_x0020_Size";

            // Document Set: In construction event
            public const string DocSetDescription = "DocumentSetDescription";
            public const string StartDate = "StartDate";
            public const string EndDate = "EndDate";
            public const string Location = "Location";
            public const string LocationDetails = "LocationDetails";
            public const string EventType = "MeetingType";
            public const string Status = "EstadoMeeting";
            public const string BodyType = "Department";
            public const string EventTool = "OnlineTool";
            public const string EventToolUrl = "UrlOnlineTool";
            public const string CalendarEventId = "OutlookMeetingId";
            public const string LastPublished = "LastPublishedDate";
            public const string Path = "Path";
            public const string Description = "Description";
            public const string OriginalDocumentId = "IdDocumentOriginal";

            public static class NotificationEntity
            {
                public const string Locale = "NotificationLocale";
                public const string Source = "NotificationSource";
            }

            //group
            //TODO prodvisional hasta crear verdaderos grupos por prodvisioning
            public const string MembersGroup = "_members";
        }

        public static class RoleGroupNames
        {
            public const string Guests = "guests";
            public const string Members = "members";
            public const string Schedulers = "schedulers";
            public const string SchedulerAssistants = "gestorschedulers";
            public const string Admins = "administrators";
            public const string MemberAssistants = "asistentemembers";
            public const string Undefined = "Undefined";
        }

        public static class AdminGroupNames
        {
            public const string Admin = "contoso-administration";
            public const string Soporte = "contoso-support";
            public const string OCP = "contoso-operations";
        }

        public static class DemoEntitiesGroupNames
        {
            public const string Admin = "DEMO-administrators";
            public const string Gestor = "DEMO-gestores";
        }

        public static class EXTERNALGroupNames
        {
            public const string Reader = "EXTERNAL-lectores";
        }

        public static class TaxonomyValuesIds
        {

            public static class EventStatus
            {
                public const string InConstruction = "1214ed8e-cfbd-4d7a-b5b7-867654823f74";
                public const string Reserved = "96909cc0-ab33-48c0-9bfd-714add3bdb4d";
                public const string Published = "8db5d28f-c8e4-463d-99bd-6e2bd4196216";
                public const string InCelebration = "7c80f44b-6586-4d98-82a8-a43acd09b99d";
                public const string Celebrated = "bd52a4c8-d276-4bb0-a1f8-12fb7adf8999";
                public const string Archived = "2522ab31-a67d-4ebd-abc4-733610aad9a9";
                public const string Cancelled = "47a651b0-921b-40ad-bf16-90d9687f4004";
            }

            public static class BodyType
            {
                public const string Conference = "63ac1500-07b4-4e25-b4ff-6fbd946389ff";
                public const string Comission = "133ead80-7a57-47ef-90d4-3eadb8b87f5f";
                public const string WorkingGroup = "639fd297-c413-445a-b5ad-18264c021f68";
            }

            public static class BodyName
            {
                public const string SampleBusinessArea = "ff0e1394-324a-4a85-ac22-6a82417a5b61"; // Comisión Sectorial, Agricultura y Desarrollo Rural
            }

            public static class TermSetTypeTask
            {
                public const string EstadoRequest = "1145f74d-26c2-4db5-ba5b-108fe66ad81f";
            }

            public static class AttendanceType
            {
                public const string Online = "8f4e0e22-2c3f-4454-8210-f4151dceeb89";

                public const string OnlineInPerson = "e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f";

                public const string InPerson = "3a16b292-5616-4ef9-8a56-2b36c20bb992";

                public const string Writtenprodcedure = "474ce74e-26c2-4ada-9dbb-516be5deb394";

                public const string DocumentationReferral = "4ab0be08-fef2-4537-ab14-3b3ae0025df3";

            }

            public static class MeetingTool
            {
                public const string Teams = "fc3ad891-8bb9-4ece-9638-26025201e131";
            }

            public static class RequestStatus
            {
                public const string Pending = "b6d26a81-48bb-4a65-a0d6-7d50b3feed1a";

                public const string Cancelled = "cb2d362d-bf4e-466b-b0b2-034cd0053291";

                public const string Rejected = "aac5817d-310f-4bfa-838a-1a1dda983bfb";

                public const string PendingDelegation = "b90d7bd3-2859-4974-a0e1-f9808701bebd";

                public const string Delegated = "fccf999f-fd9a-4370-a4a3-11400a5d19ff";

                public const string Accepted = "9628620a-c31b-4903-9bb2-b305d51821bd";

                public const string AcceptedBySystem = "de904487-484a-4250-9962-9e4a19c4a49b";

                public const string Expired = "8f52f3b1-90e6-429e-b796-6d6ad9f077b2";

                public const string PendingModification = "46707f88-14ef-456b-91f0-ae0785c89c7b";
            }

            public static class AgendaItemType
            {
                public const string ForDecision = "c83ab3b1-00a1-4023-9a2f-1697b23c13d5";
                public const string Coordination = "264eb3ec-23d4-4cc9-8ace-8a643483db0c";

                public const string Informative = "ca28d183-ce3e-4339-a516-f3aec1940622";
            }

            public static class AgreementStatus
            {
                public const string OnTable = "22abc0e5-25f1-4c07-80c5-b715dba4a1a6";

                public const string Approdved = "04012a59-69bf-4f39-99b3-799d6d9e4f26";

                public const string AgreementReject = "76426a2a-907d-4717-9857-5c704e7b1c1b";

                public const string Inprodgress = "5996ee35-33ef-4fce-820d-06db4c5cfcad";

                public const string UnanimouslyApprodved = "ddf339aa-3f6a-4a34-a24e-55b2b6efa7f5";
            }

            public static class OrderType
            {
                public const string Alphabetical = "c2c12ed4-1001-466f-977c-07249f45c3a0";
                public const string Numeric = "2cc86787-b808-4ea4-b5f7-d92f9cc9b741";
                public const string Roman = "f980cb48-6220-45d0-a9f8-3266bc31c2ba";

            }

            public static class VoteOptions
            {
                public const string Approdve = "bd558d1a-e24b-4882-865d-853c7a14289e";
                public const string Abstention = "0ae772d2-0f28-4eaf-95a9-8c857d78d51c";
                public const string Reject = "b0c9703c-16a2-4c65-966d-7f39b67c662e";
            }
        }

        public static class VoteRetestsentations
        {
            public const string TermSetId = "12e84bb5-bd3b-4640-9992-3981eafd4af4";
            public const string Scheduler = "921e8b4d-e67d-42d1-a2e8-c720737d6120";
        }

        public static class Management
        {
            public static class RequestStatus
            {
                public const string New = "New";
                public const string Inprodgress = "In prodgress";
                public const string Done = "Done";
                public const string Error = "Error";
            }

            public static class TypeOfReport
            {
                public const string Full = "Full";
                public const string Incremental = "Incremental";
            }

            public static class RequestOperations
            {
                public const string RoleManagement = "rolemngmnt";
                public const string PowerBiprocess = "powerbireport";
            }
        }

        public static class DataStorage
        {
            public static class TableNames
            {
                public const string DepartmentsEXTERNAL = "DepartmentsEXTERNAL";
                public const string RelacionBusinessAreaDivision = "RelacionBusinessAreaDivision";
                public const string DepartmentsContoso = "DepartmentsContoso";
                public const string Meetings = "Meetings";
                public const string AgendaItems = "AgendaItems";
                public const string Votes = "Votes";
                public const string Minutes = "Minutes";
                public const string DocumentsEvento = "DocumentsEvento";
                public const string TasksMinutesApprodval = "TasksMinutesApprodval";
                public const string TasksAttendance = "TasksAttendance";
                public const string TasksCertification = "TasksCertification";
                public const string TasksDelegation = "TasksDelegation";
                public const string TasksMinutesModification = "TasksMinutesModification";

                public const string DemoEntities = "DEMOEntidades";
                public const string SectionsDemoEntities = "DEMOSections";
                public const string ScopesDemoEntities = "DEMOScopes";
                public const string CancellationReasonsDemoEntities = "DEMOCancellationReasons";
                public const string DocumentsDemoEntities = "DEMODocuments";
            }

            public static class PartitionKeys
            {
                public const string DepartmentEXTERNAL = "DepartmentEXTERNAL";
                public const string AreaDivision = "AreaDivision";
                public const string DepartmentContoso = "DepartmentContoso";
                public const string Meeting = "Meeting";
                public const string AgendaItem = "AgendaItem";
                public const string Vote = "Vote";
                public const string Minutes = "Minutes";
                public const string Document = "Document";
                public const string TaskMinutes = "TaskMinutes";
                public const string TaskAttendance = "TaskAttendance";
                public const string TaskCertification = "TaskCertification";
                public const string TaskDelegation = "TaskDelegation";
                public const string TaskMinutesModification = "TaskMinutesModification";
                public const string DemoEntity = "DemoEntity";
                public const string SectionDemoEntity = "SectionDemoEntity";
                public const string ScopeDemoEntity = "ScopeDemoEntity";
                public const string CancellationReasonDemoEntity = "CancellationReasonDemoEntity";
                public const string DocumentDemoEntity = "DocumentDemoEntity";

            }
        }
    }

    public enum ComparisonOperators
    {
        BeginsWith,
        Contains,
        DateRangesOverlap,
        Eq,
        Geq,
        Gt,
        In,
        IsNotNull,
        IsNull,
        Leq,
        Lt,
        Membership,
        Neq,
        NotIncludes,
    }
}
