namespace Contoso.Portal.Model.Bodies
{

    public enum BodyType
    {
        Undefined,
        Comission,
        WorkingGroup,
        Conference
    }

    [Flags]
    public enum BodyRole
    {
        AllInvitedUsers = -1, // control notification system. Not managed by groups
        Undefined = 0, // means that the role is not controlled by the system (other user groups)
        Admin = 1,
        Scheduler = 2,
        SchedulerAssistant = 4,
        Member = 8,
        MemberAssistant = 16,
        Guest = 32,
        System = 64
    }

    [Flags]
    public enum RememberNotification
    {
        OneHour = 1,
        TwentyFourHours = 24,
    }

}