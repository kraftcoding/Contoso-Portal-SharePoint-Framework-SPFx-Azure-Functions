namespace Contoso.Portal.Data.DAO.Notifications
{
    public class NotificationMessagesJsonDAO : BaseJsonDataDAO
    {
        public override string GetFileName() => $"Notifications.json";

        public Alert Alert { get; set; }

        public Meeting Meeting { get; set; }

        public Request Request { get; set; }
        public override bool IsECNTy() => false;
    }

    public class Alert
    {

        public Alert()
        {
            Urgent = new propertiesMessage();
            ErrorGenerico = new propertiesMessage();
            AceptarDelegation = new propertiesMessage();
            InformacionDelegationAttendanceYVoto = new propertiesMessage();
            InformacionDelegationAttendance = new propertiesMessage();
            InformacionDelegationVoto = new propertiesMessage();
            RechazarDelegation = new propertiesMessage();
            MinutesValidada = new propertiesMessage();
            RechazarMinutes = new propertiesMessage();
            PublicarMinutes = new propertiesMessage();
            AceptarModificacionMinutes = new propertiesMessage();
            RechazarModificacionMinutes = new propertiesMessage();
            NuevoCertificado = new propertiesMessage();
        }

        public propertiesMessage Urgent { get; set; }

        public propertiesMessage ErrorGenerico { get; set; }

        public propertiesMessage AceptarDelegation { get; set; }

        public propertiesMessage InformacionDelegationAttendanceYVoto { get; set; }

        public propertiesMessage InformacionDelegationAttendance { get; set; }

        public propertiesMessage InformacionDelegationVoto { get; set; }

        public propertiesMessage RechazarDelegation { get; set; }

        public propertiesMessage MinutesValidada { get; set; }

        public propertiesMessage RechazarMinutes { get; set; }

        public propertiesMessage PublicarMinutes { get; set; }

        public propertiesMessage AceptarModificacionMinutes { get; set; }

        public propertiesMessage RechazarModificacionMinutes { get; set; }

        public propertiesMessage NuevoCertificado { get; set; }
    }

    public class Meeting
    {

        public Meeting()
        {
            Nueva = new propertiesMessage();
            Guardar = new propertiesMessage();
            Modificar = new propertiesMessage();
            Publicar = new propertiesMessage();
            ModificarOrdenDia = new propertiesMessage();
            ModificarCertificados = new propertiesMessage();
            ModificarMinutes = new propertiesMessage();
            EnCelebracion = new propertiesMessage();
            Celebrada = new propertiesMessage();
            Finalizada = new propertiesMessage();
            Cancelada = new propertiesMessage();
        }

        public propertiesMessage Nueva { get; set; }
        public propertiesMessage Guardar { get; set; }
        public propertiesMessage Modificar { get; set; }
        public propertiesMessage ModificarOrdenDia { get; set; }
        public propertiesMessage ModificarCertificados { get; set; }
        public propertiesMessage ModificarMinutes { get; set; }
        public propertiesMessage Publicar { get; set; }
        public propertiesMessage EnCelebracion { get; set; }
        public propertiesMessage Celebrada { get; set; }
        public propertiesMessage Finalizada { get; set; }
        public propertiesMessage Archivada { get; set; }
        public propertiesMessage Cancelada { get; set; }
        public propertiesMessage Reminder1H { get; set; }
        public propertiesMessage Reminder24H { get; set; }
    }

    public class Request
    {

        public Request()
        {
            Attendance = new propertiesMessage();
            DelegarAttendance = new propertiesMessage();
            DelegateVote = new propertiesMessage();
            GeneracionCertificado = new propertiesMessage();
            ModificacionMinutes = new propertiesMessage();
        }

        public propertiesMessage Attendance { get; set; }

        public propertiesMessage DelegarAttendance { get; set; }

        public propertiesMessage DelegateVote { get; set; }

        public propertiesMessage GeneracionCertificado { get; set; }

        public propertiesMessage ModificacionMinutes { get; set; }
    }

    public class propertiesMessage
    {
        public String Type { get; set; }

        public String Priority { get; set; }

        public Boolean SendMail { get; set; }

        public Locates Title { get; set; }

        public Locates Body { get; set; }
    }

    public class Locates
    {
        public string esES { get; set; } = string.ECNTy;
        public string caES { get; set; } = string.ECNTy;
        public string euES { get; set; } = string.ECNTy;

        public string glES { get; set; } = string.ECNTy;
    }
}