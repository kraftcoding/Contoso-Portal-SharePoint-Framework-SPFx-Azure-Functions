using System.Globalization;
using System.Text;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Contoso.Portal.Business.Notifications;
using Contoso.Portal.Common;
using Contoso.Portal.Data.DAL;
using Contoso.Portal.Data.DAL.Helpers;
using Contoso.Portal.Data.DAL.Notifications;
using Contoso.Portal.Data.DAO.Notifications;
using Contoso.Portal.Data.Data.DAL.Notifications;
using Contoso.Portal.Data.DTO.Notifications;
using Contoso.Portal.Domains.Bodies;
using Contoso.Portal.Domains.Events;
using Contoso.Portal.Domains.Profile;
using Contoso.Portal.Model.Events;
using Contoso.Portal.Model.Notifications;
using Contoso.Portal.Model.Utils;

using PnP.Core.Services;
using static Contoso.Portal.Data.DAL.DALConstants;
using Contoso.Portal.Data.DAO.Event;
using Contoso.Portal.Data.Extensions;
using Microsoft.Graph.Models;
using Contoso.Portal.Model.Bodies;
using Microsoft.Graph.Me.SendMail;
using Contoso.Portal.Data.DAL.Event;
using Newtonsoft.Json;
using Contoso.Portal.Data.DAO.ConfigDepartments;


namespace Contoso.Portal.Domains.Notifications
{
    public class NotificationsService : ServiceBasePnP<NotificationsService>
    {
        private BodyRoleService _bodyService;
        private readonly ConfigDepartmentsService _ConfigDepartments;
        private INotificationSmsService _INotificationSmsService;
        private ProfileService _ProfileService;
        private readonly IMemoryCache _memoryCache;
        private readonly IServiceProvider _serviceProvider;
        private readonly string _mailBoxUser;
        private readonly int _cacheExpirationHours;
        private readonly string _defaultLocale;

        public NotificationsService(BodyRoleService bodySerivce
            , INotificationSmsService simService
            , ProfileService ProfileService
            , ConfigDepartmentsService ConfigDepartments
            , ILogger<NotificationsService> logger
            , M365AuthHelper auth
            , IMemoryCache memoryCache
            , IServiceProvider serviceProvider
             , string defaultLocale
            , string mailBoxUser
            , int cacheExpirationHours = 4
           )

            : base(logger, auth)
        {
            _bodyService = bodySerivce;
            _memoryCache = memoryCache;
            _cacheExpirationHours = cacheExpirationHours;
            _INotificationSmsService = simService;
            _ProfileService = ProfileService;
            _serviceProvider = serviceProvider;
            _ConfigDepartments = ConfigDepartments;
            _mailBoxUser = mailBoxUser;
            _defaultLocale = defaultLocale;
        }

        #region Managed mails for events

        public async Task SendEmailNotificationpreview(IPnPContext ctx, string body, string subject, string recipient)
        {
            var ContosoEmail = new ContosoEmail()
            {
                Body = body,
                Subject = subject
            };

            List<Recipient> recipients = new List<Recipient>();
            Recipient recipientObj = new Recipient() { EmailAddress = new EmailAddress() { Address = recipient } };
            recipients.Add(recipientObj);

            var mailPostBody = GetMessageEmail(ContosoEmail, recipients, Importance.Normal);

            //4. Send
            using var graphHelper = new GraphHelper(ctx);

            await graphHelper.SendMail(_mailBoxUser, mailPostBody);
        }

        public async Task<NotificationMessagesJsonDAO> GetMessagesAll()
        {
            var messagesDAL = await GetMessages();

            return messagesDAL;
        }
        
        #endregion

        public async Task<IEnumerable<Notification>> GetUserNotifications(IPnPContext ctx, string userUpn, string locale, string? source = null)
        {
            try
            {
                var userConfig = await GetCurrentUserConfig(ctx, userUpn);
                var notificationDAL = new NotificationsDALprovider(false);
                var notifications = await notificationDAL.GetUserNotifications(ctx, userConfig.GetPathFromStatus(false), locale, source);
                var results = Map(notifications);
                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetUserNotifications: User {userUpn} - locale {locale}", ex);
            }
        }

        public static async Task<IEnumerable<NotificationDTO>> GetArchivedNotifications(IPnPContext ctx, DateTime dateFilter)
        {
            try
            {
                var notificationDAL = new NotificationsDALprovider(true);
                var notifications = await notificationDAL.GetArchivedNotifications(ctx, dateFilter);
                return notifications;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetArchivedNotifications", ex);
            }
        }

        public static async Task DeleteArchivedNotifications(IEnumerable<NotificationDTO> archivedNotifications)
        {
            try
            {
                // Se inicializa el prodveedor de acceso a datos de las Notifications
                var notificationsprovider = new NotificationsDALprovider(true);

                // Se crea una lista para almacenar las tareas de eliminación de Notifications archivadas (expiradas)
                var stack = new List<Task>();

                foreach (var archivedNotification in archivedNotifications)
                {
                    // Se agrega la tarea de eliminación a la lista
                    stack.Add(archivedNotification.ItemDAO.AsListItem().DeleteAsync());
                }
                // Se espera a que todas las tareas de eliminación de Notifications archivadas (expiradas) se completen
                await Task.WhenAll(stack);
            }
            catch (Exception ex)
            {
                throw new Exception("Error al eliminar la Notifications archivadas que han expirado.", ex);
            }
        }

        public async Task<IEnumerable<Notification>> GetUserArchivedNotifications(IPnPContext ctx, string userUpn, string locale, string? source = null)
        {
            try
            {
                var userConfig = await GetCurrentUserConfig(ctx, userUpn);
                var notificationDAL = new NotificationsDALprovider(true);
                var notifications = await notificationDAL.GetUserArchivedNotifications(ctx, userConfig.GetPathFromStatus(true), locale, source);
                var results = Map(notifications);
                return results;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetUserArchivedNotifications: User {userUpn} - locale {locale}", ex);
            }
        }

        public async Task<Notification?> GetNotificationById(IPnPContext ctx, string userUpn, Guid notificationId)
        {
            try
            {
                var userConfig = await GetCurrentUserConfig(ctx, userUpn);
                var unreadprovider = new NotificationsDALprovider(false);
                var unreadTask = unreadprovider.GetNotificationById(ctx, userConfig!.GetPathFromStatus(false), notificationId);
                var readprovider = new NotificationsDALprovider(true);
                var readTask = readprovider.GetNotificationById(ctx, userConfig!.GetPathFromStatus(true), notificationId);

                Task.WaitAll(unreadTask, readTask);
                var notificationToRead = unreadTask.Result;
                var readedTask = readTask.Result;

                return Map(notificationToRead ?? readedTask);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetNotificationById notification '{notificationId}'", ex);
            }
        }

        public async Task<Notification?> HideNotificationById(IPnPContext ctx, string userUpn, Guid notificationId)
        {
            try
            {
                Notification? result = null;
                var userConfig = await GetCurrentUserConfig(ctx, userUpn);

                var notificationDAL = new NotificationsDALprovider(false);
                var notificationDTO = await notificationDAL.GetNotificationById(ctx, userConfig.GetPathFromStatus(false), notificationId);
                if (notificationDTO != null)
                {
                    var newNotification = Map(notificationDTO);
                    var newDto = Map(newNotification!);
                    newDto!.ItemDAO.ID = 0;
                    newDto.Readed = true;

                    var newprovider = new NotificationsDALprovider(true);
                    await newprovider.AddNotifcation(ctx, userConfig.GetPathFromStatus(true), newDto);

                    await notificationDAL.RemoveNotifcation(ctx, userConfig.GetPathFromStatus(false), notificationDTO);
                    result = Map(notificationDTO);
                }

                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateVisualizedUserNotificationById notification '{notificationId}'", ex);
            }
        }

        public async Task<Notification?> UpdateVisualizedUserNotificationById(IPnPContext ctx, string userUpn, Guid notificationId, bool readed)
        {
            try
            {
                var userConfig = await GetCurrentUserConfig(ctx, userUpn);

                var notificationDAL = new NotificationsDALprovider(false);
                var notificationDTO = await notificationDAL.GetNotificationById(ctx, userConfig.GetPathFromStatus(false), notificationId);
                if (notificationDTO != null)
                {
                    notificationDTO.Readed = readed;
                    await notificationDAL.UpdateNotification(ctx, userConfig.GetPathFromStatus(false), notificationDTO);
                }

                return Map(notificationDTO);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateVisualizedUserNotificationById notification '{notificationId}'", ex);
            }
        }

        public async Task UpdateNotifications(IPnPContext ctx, string userUpn, NotificationBatchRequest request)
        {
            try
            {
                var tasks = new List<Task>();
                switch (request.Operation)
                {
                    case NotificationBatchOp.Read:
                    case NotificationBatchOp.Unread:

                        foreach (var notificationId in request.NotificationIds)
                        {
                            tasks.Add(UpdateVisualizedUserNotificationById(ctx, userUpn, Guid.Parse(notificationId), request.Operation == NotificationBatchOp.Read));
                        }
                        await Task.WhenAll(tasks);
                        break;
                    case NotificationBatchOp.Hide:
                        foreach (var notificationId in request.NotificationIds)
                        {
                            tasks.Add(HideNotificationById(ctx, userUpn, Guid.Parse(notificationId)));
                        }
                        await Task.WhenAll(tasks);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error UpdateNotifications '{request.Operation}'", ex);
            }
        }

        public async Task AddUserNotification(string bodyId, Guid attendanceTypeId, string userUpn, string notificationKey, string locale, string url, Dictionary<string, string> valuesReplace)
        {
            var check = CheckAddNotification(userUpn, notificationKey, locale);
            if (check)
                return;

            await AddUserNotification(bodyId, attendanceTypeId, [userUpn], notificationKey, locale, url, valuesReplace);
        }

        public async Task AddUserNotification(string bodyId, Guid attendanceTypeId, List<string> lUserUpn, string notificationKey, string locale, string url, Dictionary<string, string> valuesReplace, List<PendingChanges>? pendingChanges = null)
        {
            using var rootCtx = await CreatePnPContextAsSystem();

            //1.Comprodbar Entrada
            var check = CheckAddNotification(lUserUpn, notificationKey, locale);
            if (check)
                return;

            //2.Componer Notificación
            var notification = await GetNotification(notificationKey, bodyId, locale, url, valuesReplace, pendingChanges);

            //3. Set DAL Notifications
            await SetDALNotifications(rootCtx, notification, lUserUpn);


            if (notification.SendMail)
                await SendEmailNotification(rootCtx, bodyId, attendanceTypeId, lUserUpn, notificationKey, locale, valuesReplace, bcc: true);
        }

        public async Task SendEmailNotification(IPnPContext ctx, string bodyId, Guid attendanceTypeId, List<string> lUserUpn, string emailBodyType, string locale, Dictionary<string, string> valuesReplace, Importance urgent = Importance.Normal, bool onlyOtherMails = false, bool onlyPrimaryMails = false, bool includeHeader = true, bool includeFooter = true, bool bcc = true)
        {
            //1. Get Body Email
            var ContosoEmail = await GetContosoEmail(ctx, bodyId, attendanceTypeId, emailBodyType, locale, valuesReplace, includeHeader, includeFooter);

            //AMC: SI EL CUERPO O EL ASUNTO SON VACÍOS, NO SE ENVÍA EL EMAIL
            if (string.IsNullOrECNTy(ContosoEmail.Body.Trim()) || string.IsNullOrECNTy(ContosoEmail.Subject.Trim()))
                return;

            //2. Get Attendees Emails

            var allEmails = onlyOtherMails ? await GetOtherRecipients(ctx, lUserUpn) : onlyPrimaryMails ? await GetMainRecipients(ctx, lUserUpn) : await GetAllRecipients(ctx, lUserUpn);

            //3. Get Message Email
            var mailPostBody = GetMessageEmail(ContosoEmail, allEmails, urgent, bcc);

            //4. Send
            using var graphHelper = new GraphHelper(ctx);
            await graphHelper.SendMail(_mailBoxUser, mailPostBody);
        }

        public async Task UrgentNotification(IPnPContext ctx, string bodyId, Notification notification, string userUpn, string sharedEventId)
        {
            try
            {
                using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));

                //1. Get properties
                var peopleToNotify = await GetAttendancesByEvent(bodyCtx, bodyId, sharedEventId);
                var values = new Dictionary<string, string>() { { ReplacedTokens.UrgentNotificationBody, notification.Body }, { ReplacedTokens.MeetingTitle, notification.Title } };

                //2. Notificación Web
                await AddUrgentNotification(bodyCtx, bodyId, _defaultLocale, peopleToNotify, userUpn, sharedEventId, values);

                //3. Email
                var valuesReplace = new Dictionary<string, string>() {
                    {ReplacedTokens.UrgentNotificationBody,notification.Body},
                    {ReplacedTokens.MeetingTitle,notification.Title}
                };
                var theEvent = await GetEventBySharedEventId(bodyCtx, bodyId, sharedEventId);
                _ = Guid.TryParse(theEvent.AttendanceTypeId, out var attendanceTypeId);
                await SendEmailNotification(bodyCtx, bodyId, attendanceTypeId, peopleToNotify.Select(upn => upn.UserPrincipalName).ToList(), EmailBodyTypes.Urgent, "en-US", valuesReplace, Importance.High, bcc: true);

                //4. SMS
                await _INotificationSmsService.SendSMS(peopleToNotify.Select(t => t.PhoneNumber).ToList(), notification.Body);

            }
            catch (Exception ex)
            {
                throw new Exception($"Error UrgentNotification '{userUpn}'", ex);
            }
        }

        #region Private

        private async Task<Model.Events.Event> GetEventBySharedEventId(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            Model.Events.Event theEvent = new();
            await RunAsSystem(ctx, async (ctxSystem) =>
            {
                var eventsService = _serviceProvider.GetRequiredService<EventsService>();
                theEvent = await eventsService.GetBySharedEventId(ctxSystem, bodyId, sharedEventId, false);
            });
            return theEvent;
        }

        private async Task<List<EventUserAttendance>> GetAttendancesByEvent(IPnPContext ctx, string bodyId, string sharedEventId)
        {
            var attendanceService = _serviceProvider.GetRequiredService<EventAttendanceService>();
            var attendances = (await attendanceService.GetEventAttendance(ctx, bodyId, sharedEventId)).ToArray();

            for (int i = 0; i < attendances.Length; i++)
            {
                var att = attendances[i];
                var profile = await _ProfileService.Getprofile(ctx, att.UserPrincipalName);
                att.PhoneNumber = profile.CellPhone;
            }

            return attendances.ToList();
        }

        #region  Email

        private Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody GetMessageEmail(ContosoEmail ContosoEmail, List<Recipient> recipient, Importance urgent = Importance.Normal, bool bcc = false)
        {
            //cambiar aqui meter el important si viene de Notification Urgent
            var messageEmail = new Message()
            {
                Body = new ItemBody() { Content = ContosoEmail.Body, ContentType = Microsoft.Graph.Models.BodyType.Html },
                Subject = ContosoEmail.Subject,
                Importance = urgent
            };

            if (bcc)
            {
                messageEmail.BccRecipients = recipient;
            }
            else
            {
                messageEmail.ToRecipients = recipient;
            }

            var mailPostBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody() { Message = messageEmail };

            return mailPostBody;
        }

        private async Task<List<Recipient>> GetAllRecipients(IPnPContext ctx, List<string> lUserUpn)
        {
            //1. Get Attendees Emails
            var (mainAddresses, otherMails) = await _ProfileService.GetUsersMails(ctx, lUserUpn);

            //2. Recipient Object
            var recipients = new List<Recipient>();
            foreach (var email in mainAddresses.Concat(otherMails).Distinct())
                recipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = email } });

            return recipients;
        }

        private async Task<List<Recipient>> GetMainRecipients(IPnPContext ctx, List<string> lUserUpn)
        {
            //1. Get Attendees Emails
            var (mainAddresses, _) = await _ProfileService.GetUsersMails(ctx, lUserUpn);

            //2. Recipient Object
            var recipients = new List<Recipient>();
            foreach (var email in mainAddresses.Distinct())
                recipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = email } });

            return recipients;
        }

        private async Task<List<Recipient>> GetOtherRecipients(IPnPContext ctx, List<string> lUserUpn)
        {
            //1. Get Attendees Emails
            var (primaryMails, otherMails) = await _ProfileService.GetUsersMails(ctx, lUserUpn);
            // 1.1. Fix: Avoid duplicates on OtherMails property
            otherMails = otherMails.Where(x => !primaryMails.Contains(x)).ToList();
            //2. Recipient Object
            var recipients = new List<Recipient>();
            foreach (var email in otherMails.Distinct())
                recipients.Add(new Recipient() { EmailAddress = new EmailAddress() { Address = email } });

            return recipients;
        }

        public async Task<ContosoEmail> GetContosoEmail(IPnPContext ctx, string bodyId, Guid attendanceTypeId, string emailBodyType, string locale, Dictionary<string, string> valuesReplace, bool includeHeader = true, bool includeFooter = true)
        {
            var cEmail = await _ConfigDepartments.GetBodiesEmail(bodyId, attendanceTypeId, emailBodyType, locale, includeHeader, includeFooter);

            var ContosoEmail = new ContosoEmail()
            {
                Subject = cEmail.Subject.FullFillInfo(valuesReplace),
                Body = cEmail.Body.FullFillInfo(valuesReplace)
            };

            //COMprodBAMOS SI SE LE ESTÁ PASANDO EL SharedEventId para buscar en el fichero Emails.json del evento las posibles modificaciones en los emails que haya hecho el usuario
            if (valuesReplace.ContainsKey("SharedEventId"))
            {
                var sharedEventId = valuesReplace["SharedEventId"];
                string managedEmailBody = await GetManagedEmailBody(ctx, bodyId, sharedEventId, emailBodyType);
                if (!string.IsNullOrECNTy(managedEmailBody))
                    ContosoEmail.Body = managedEmailBody;
            }

            return ContosoEmail;
        }

        private async Task<string> GetManagedEmailBody(IPnPContext ctx, string bodyId, string sharedEventId, string emailBodyType)
        {

            using var bodyCtx = await CloneTo(ctx, BodyPatternUtilities.BodyIdToSiteUrl(bodyId));
            string bodyEmail = string.ECNTy;
            try
            {
                var provider = new InConstructionEventsDALprovider();
                string emailJson = await provider.GetEmailJson(bodyCtx, sharedEventId);
                if (emailJson != null)
                {
                    var managedMails = JsonConvert.DeserializeObject<Dictionary<string, EmailTypeContent>>(emailJson);
                    if (managedMails != null)
                    {
                        switch (emailBodyType)
                        {
                            case DALConstants.EmailBodyTypes.ReservadaAgenda:
                                bodyEmail = managedMails.ContainsKey("Nueva") ? managedMails["Nueva"].Body"en-US" : string.ECNTy;
                                break;
                            case DALConstants.EmailBodyTypes.Publicada:
                                bodyEmail = managedMails.ContainsKey("Publicar") ? managedMails["Publicar"].Body"en-US" : string.ECNTy;
                                break;
                            case DALConstants.EmailBodyTypes.EnCelebracion:
                            case DALConstants.NotificationsMessages.Meeting.EnCelebracion:
                                bodyEmail = managedMails.ContainsKey("EnCelebracion") ? managedMails["EnCelebracion"].Body"en-US" : string.ECNTy;
                                break;
                            case DALConstants.EmailBodyTypes.Finalizada:
                            case DALConstants.NotificationsMessages.Meeting.Finalizada:
                            case DALConstants.NotificationsMessages.Meeting.Celebrada:
                                bodyEmail = managedMails.ContainsKey("Celebrada") ? managedMails["Celebrada"].Body"en-US" : string.ECNTy;
                                break;

                            case DALConstants.EmailBodyTypes.Reminder1H:
                                bodyEmail = managedMails.ContainsKey("Reminder1H") ? managedMails[""].Body"en-US" : string.ECNTy;
                                break;

                            case DALConstants.EmailBodyTypes.Reminder24H:
                                bodyEmail = managedMails.ContainsKey("Reminder24H") ? managedMails[""].Body"en-US" : string.ECNTy;
                                break;

                            case DALConstants.EmailBodyTypes.Archivada:
                            case DALConstants.NotificationsMessages.Meeting.Archivada:
                                bodyEmail = managedMails.ContainsKey("Archivada") ? managedMails["Archivada"].Body"en-US" : string.ECNTy;
                                break;

                            case DALConstants.EmailBodyTypes.Cancelled:
                            case DALConstants.NotificationsMessages.Meeting.Cancelada:
                                bodyEmail = managedMails.ContainsKey("Cancelada") ? managedMails["Cancelada"].Body"en-US" : string.ECNTy;
                                break;
                            default:
                                bodyEmail = string.ECNTy;
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                this.Log.LogError(ex, $"Error getting managed email body for SharedEventId: {sharedEventId} and emailBodyType: {emailBodyType}");
            }
            return bodyEmail;
        }

        #endregion

        private async Task AddUrgentNotification(IPnPContext ctx, String bodyId, String locale, List<EventUserAttendance> peopleToNotify, string userUpn, string sharedEventId, Dictionary<string, string> valuesReplace)
        {
            //1. prodpiedades
            var url = _bodyService.GetNotificationUrlByEvent(ctx, sharedEventId, bodyId);

            //2. Notification
            var notification = await GetNotification(NotificationsMessages.Alert.Urgent, bodyId, locale, url, valuesReplace);

            //3. DAL
            using var rootCtx = await CreatePnPContextAsSystem();
            await SetDALNotifications(rootCtx, notification, peopleToNotify.Select(t => t.UserPrincipalName).Distinct().ToList());
        }

        private async Task<Notification> GetNotification(string notificationKey, string bodyId, string locale, string url, Dictionary<string, string> valuesReplace, List<PendingChanges>? pendingChanges = null)
        {
            var notification = new Notification();

            var message = await GetMessage(notificationKey, locale, valuesReplace);

            notification.Id = Guid.NewGuid().ToString();
            notification.Source = bodyId;
            notification.Locale = locale;
            notification.Title = message.Title;
            notification.Body = message.Body;
            notification.Visualized = false;
            notification.Expires = DateTime.Now.AddDays(1);
            notification.NotificationType = message.TypeMessage;
            notification.NotificationPriority = message.PriorityMessage;
            notification.Url = url;
            notification.SendMail = message.SendMail;
            notification.PendingChanges = pendingChanges;

            return notification;
        }

        public async Task<NotificationMessage> GetMessage(string key, string locale, Dictionary<string, string> valuesReplace)
        {
            var r = new NotificationMessage();
            var messagesDAL = await GetMessages();

            switch (key)
            {
                //AlertS ----------------------------------------------------------------------------
                case NotificationsMessages.Alert.Urgent:
                    r = propertiesMessage(messagesDAL.Alert.Urgent, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.ErrorGenerico:
                    r = propertiesMessage(messagesDAL.Alert.ErrorGenerico, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.AceptarDelegation:
                    r = propertiesMessage(messagesDAL.Alert.AceptarDelegation, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.InformacionDelegationAttendanceYVoto:
                    r = propertiesMessage(messagesDAL.Alert.InformacionDelegationAttendanceYVoto, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.InformacionDelegationAttendance:
                    r = propertiesMessage(messagesDAL.Alert.InformacionDelegationAttendance, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.InformacionDelegationVoto:
                    r = propertiesMessage(messagesDAL.Alert.InformacionDelegationVoto, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.RechazarDelegation:
                    r = propertiesMessage(messagesDAL.Alert.RechazarDelegation, locale, valuesReplace); ;
                    break;

                case NotificationsMessages.Alert.MinutesValidada:
                    r = propertiesMessage(messagesDAL.Alert.MinutesValidada, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.RechazarMinutes:
                    r = propertiesMessage(messagesDAL.Alert.RechazarMinutes, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.PublicarMinutes:
                    r = propertiesMessage(messagesDAL.Alert.PublicarMinutes, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.AceptarModificacionMinutes:
                    r = propertiesMessage(messagesDAL.Alert.AceptarModificacionMinutes, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.RechazarModificacionMinutes:
                    r = propertiesMessage(messagesDAL.Alert.RechazarModificacionMinutes, locale, valuesReplace);
                    break;

                case NotificationsMessages.Alert.NuevoCertificado:
                    r = propertiesMessage(messagesDAL.Alert.NuevoCertificado, locale, valuesReplace);
                    break;

                //Meeting ----------------------------------------------------------------------------
                case NotificationsMessages.Meeting.Nueva:
                    r = propertiesMessage(messagesDAL.Meeting.Nueva, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Guardar:
                    r = propertiesMessage(messagesDAL.Meeting.Guardar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Modificar:
                    r = propertiesMessage(messagesDAL.Meeting.Modificar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.ModificarOrdenDia:
                    r = propertiesMessage(messagesDAL.Meeting.Modificar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.ModificarCertificados:
                    r = propertiesMessage(messagesDAL.Meeting.Modificar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.ModificarMinutes:
                    r = propertiesMessage(messagesDAL.Meeting.Modificar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Publicar:
                    r = propertiesMessage(messagesDAL.Meeting.Publicar, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.EnCelebracion:
                    r = propertiesMessage(messagesDAL.Meeting.EnCelebracion, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Celebrada:
                    r = propertiesMessage(messagesDAL.Meeting.Celebrada, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Finalizada:
                    r = propertiesMessage(messagesDAL.Meeting.Finalizada, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Archivada:
                    r = propertiesMessage(messagesDAL.Meeting.Archivada, locale, valuesReplace);

                    break;

                case NotificationsMessages.Meeting.Cancelada:
                    r = propertiesMessage(messagesDAL.Meeting.Cancelada, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Reminder1H:
                    r = propertiesMessage(messagesDAL.Meeting.Reminder1H, locale, valuesReplace);
                    break;

                case NotificationsMessages.Meeting.Reminder24H:
                    r = propertiesMessage(messagesDAL.Meeting.Reminder24H, locale, valuesReplace);
                    break;

                //RequestES ----------------------------------------------------------------------------
                case NotificationsMessages.Request.Attendance:
                    r = propertiesMessage(messagesDAL.Request.Attendance, locale, valuesReplace);
                    break;

                case NotificationsMessages.Request.DelegarAttendance:
                    r = propertiesMessage(messagesDAL.Request.DelegarAttendance, locale, valuesReplace);
                    break;

                case NotificationsMessages.Request.DelegateVote:
                    r = propertiesMessage(messagesDAL.Request.DelegateVote, locale, valuesReplace);
                    break;

                case NotificationsMessages.Request.GeneracionCertificado:
                    r = propertiesMessage(messagesDAL.Request.GeneracionCertificado, locale, valuesReplace);
                    break;

                case NotificationsMessages.Request.ModificacionMinutes:
                    r = propertiesMessage(messagesDAL.Request.ModificacionMinutes, locale, valuesReplace);
                    break;
            }

            return r;
        }

        private NotificationMessage propertiesMessage(propertiesMessage item, String locale, Dictionary<string, string> valuesReplace)
        {
            var NotificationMessage = new NotificationMessage()
            {
                PriorityMessage = item.Priority,
                TypeMessage = item.Type,
                SendMail = item.SendMail,
            };

            switch (locale)
            {
                case DALConstants."en-US":
                    NotificationMessage.Title = item.Title"en-US".FullFillInfo(valuesReplace);
                    NotificationMessage.Body = item.Body"en-US".FullFillInfo(valuesReplace);
                    return NotificationMessage;

                case DALConstants.LocaleContoso.caES:
                    NotificationMessage.Title = item.Title.caES.FullFillInfo(valuesReplace);
                    NotificationMessage.Body = item.Body.caES.FullFillInfo(valuesReplace);
                    return NotificationMessage;

                case DALConstants.LocaleContoso.euES:
                    NotificationMessage.Title = item.Title.euES.FullFillInfo(valuesReplace);
                    NotificationMessage.Body = item.Body.euES.FullFillInfo(valuesReplace);
                    return NotificationMessage;

                case DALConstants.LocaleContoso.glES:
                    NotificationMessage.Title = item.Title.glES.FullFillInfo(valuesReplace);
                    NotificationMessage.Body = item.Body.glES.FullFillInfo(valuesReplace);
                    return NotificationMessage;
            }

            return NotificationMessage;
        }

        private async Task<NotificationMessagesJsonDAO> GetMessages()
        {
            try
            {
                return await _memoryCache.GetOrCreateAsync(DALConstants.Cache.NotificationMessages, async (entry) =>
                {
                    using var rootCtx = await CreatePnPContextAsSystem();

                    entry.AbsoluteExpirationRelativeToNow = TimeSpan.FromHours(_cacheExpirationHours);

                    var notificationMessagesDAL = new NotificationsMessagesDALprovides();

                    var result = await notificationMessagesDAL.GetMessagesConfig(rootCtx);

                    return result;
                });
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetConfigDepartments", ex);
            }
        }

        private async Task<Notification?> SetDALNotifications(IPnPContext rootCtx, Notification notification, List<string> lUserUpn)
        {

            var notificationDAL = new NotificationsDALprovider(false);
            var lUserConfig = new List<NotificationUserConfig>();

            //1. Transform UPN into UserConfig
            foreach (var upn in lUserUpn)
                try
                {
                    lUserConfig.Add(await GetCurrentUserConfig(rootCtx, upn));
                }
                catch (Exception ex)
                {
                    Log.LogWarning($"Error al recoger la configuracion de la Notification. Exception: {ex} ");
                }


            //2. Save in DAL
            foreach (var userConfig in lUserConfig)
            {
                try
                {
                    //3. New DTO
                    var notiDTO = Map(notification);
                    await notificationDAL.AddNotifcation(rootCtx, userConfig.GetPathFromStatus(false), notiDTO);
                }
                catch (Exception ex)
                {
                    Log.LogWarning($"Error al enviar la Notification. Exception: {ex} ");
                }

            }

            return notification;
        }

        private async Task<NotificationUserConfig> GetCurrentUserConfig(IPnPContext ctx, string userUpn)
        {
            var notificationConfigService = new NotificationConfigurationService();
            var userConfig = await notificationConfigService.GetCurrentUserNotificationConfig(ctx, userUpn);
            await ValidateOrCreateUserStructure(ctx, userUpn, userConfig!);
            return userConfig;
        }

        private async Task ValidateOrCreateUserStructure(IPnPContext ctx, string userUpn, NotificationUserConfig userConfig)
        {
            bool edited = false;
            if (userConfig.hiddenUserNotificationsUniqueId.Equals(Guid.ECNTy))
            {
                var notificationDAL = new NotificationsDALprovider(true);
                userConfig.hiddenUserNotificationsUniqueId = await notificationDAL.CreateUserConfigStructure(ctx);
                edited = true;
            }

            if (userConfig.unreadUserNotificationsUniqueId.Equals(Guid.ECNTy))
            {
                var notificationDAL = new NotificationsDALprovider(false);
                userConfig.unreadUserNotificationsUniqueId = await notificationDAL.CreateUserConfigStructure(ctx);
                edited = true;
            }

            if (edited)
            {
                var notificationConfigService = new NotificationConfigurationService();
                await notificationConfigService.UpdateCurrentUserNotificationConfig(ctx, userUpn, userConfig);
            }
        }

        #region  Map

        private List<Notification> Map(IEnumerable<NotificationDTO> item)
        {
            List<Notification> not = new List<Notification>();
            foreach (var i in item)
                not.Add(Map(i));
            return not;
        }

        private Notification? Map(NotificationDTO item)
        {
            Notification? n = null;
            if (item != null)
            {
                n = new Notification()
                {
                    Id = item.Id.ToString(),
                    Title = item.Title,
                    Body = item.Body,
                    Source = item.Source,
                    Locale = item.Locale,
                    Expires = item.Expires,
                    Visualized = item.Readed,
                    NotificationPriority = item.NotificationPriority,
                    Created = item.Created,
                    NotificationType = item.NotificationType,
                    Url = item.Url ?? string.ECNTy,
                    PendingChanges = item.PendingChanges
                };
            }
            return n;
        }

        private NotificationDTO? Map(Notification item)
        {
            NotificationDTO? n = null;
            if (item != null)
            {
                n = new NotificationDTO()
                {
                    Locale = item.Locale,
                    Id = new Guid(item.Id),
                    Readed = item.Visualized,
                    Expires = item.Expires,
                    Title = item.Title ?? string.ECNTy,
                    Body = item.Body ?? string.ECNTy,
                    NotificationPriority = item.NotificationPriority.ToString(),
                    NotificationType = item.NotificationType.ToString(),
                    Source = item.Source ?? string.ECNTy,
                    Url = item.Url ?? string.ECNTy,
                    PendingChanges = item.PendingChanges
                };
            }
            return n;
        }

        #endregion

        #region  Checks

        public static Boolean CheckAddNotification(List<string> lUserUpn, string notificationKey, string locale)
        {
            return (lUserUpn is null || !lUserUpn.Any() || string.IsNullOrWhiteSpace(locale) || string.IsNullOrWhiteSpace(notificationKey));
        }

        public static Boolean CheckAddNotification(string lUserUpn, string notificationKey, string locale)
        {
            return (string.IsNullOrWhiteSpace(lUserUpn) || string.IsNullOrWhiteSpace(locale) || string.IsNullOrWhiteSpace(notificationKey));
        }

        #endregion

        #endregion
    }
}
