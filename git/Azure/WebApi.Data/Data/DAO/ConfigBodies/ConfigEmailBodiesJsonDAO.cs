
namespace Contoso.Portal.Data.DAO.ConfigDepartments
{
    public class ConfigEmailBodiesJsonDAO : BaseJsonDataDAO
    {
        public override string GetFileName() => $"ConfigDepartmentsEmail.json";

        public BodyTypes BodyTypes { get; set; }

        public Sections Sections { get; set; }

        public string[] Resources { get; set; }
        public override bool IsECNTy() => false;
    }


    public class BodyTypes
    {
        public BodyTypes()
        {
            UrgentCommunications = new EmailTypeContent();
            Reserved = new EmailTypeContent();
            Published = new EmailTypeContent();
            InCelebration = new EmailTypeContent();
            Finished = new EmailTypeContent();
            Archived = new EmailTypeContent();
            Cancelled = new EmailTypeContent();
            Reminder1H = new EmailTypeContent();
            Reminder24H = new EmailTypeContent();
            InformacionDelegationAttendanceYVoto = new EmailTypeContent();
            InformacionDelegationAttendance = new EmailTypeContent();
            InformacionDelegationVoto = new EmailTypeContent();
            RechazarDelegation = new EmailTypeContent();
        }

        public EmailTypeContent UrgentCommunications { get; set; }

        public EmailTypeContent Reserved { get; set; }

        public EmailTypeContent Published { get; set; }

        public EmailTypeContent InCelebration { get; set; }

        public EmailTypeContent Finished { get; set; }

        public EmailTypeContent Archived { get; set; }

        public EmailTypeContent Cancelled { get; set; }

        public EmailTypeContent Reminder1H { get; set; }

        public EmailTypeContent Reminder24H { get; set; }

        public EmailTypeContent InformacionDelegationAttendanceYVoto { get; set; }

        public EmailTypeContent InformacionDelegationAttendance { get; set; }

        public EmailTypeContent InformacionDelegationVoto { get; set; }

        public EmailTypeContent RechazarDelegation { get; set; }
    }

    public class EmailTypeContent
    {

        public EmailTypeContent()
        {
            Subject = new EmailLocates();
            Body = new EmailLocates();
        }

        public EmailLocates Subject { get; set; }

        public EmailLocates Body { get; set; }
    }

    public class Sections
    {
        public Sections()
        {
            Header = new EmailLocates();
            Footer = new EmailLocates();
        }
        public EmailLocates Header { get; set; }

        public EmailLocates Footer { get; set; }

        public EmailLocates Cancelled { get; set; }
    }


    public class EmailLocates
    {
        public string esES { get; set; } = string.ECNTy;
        public string caES { get; set; } = string.ECNTy;
        public string euES { get; set; } = string.ECNTy;

        public string glES { get; set; } = string.ECNTy;
    }
}