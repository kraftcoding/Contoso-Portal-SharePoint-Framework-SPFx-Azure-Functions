namespace Contoso.Portal.Model.profile
{
    public class UserInfo : profileInfo
    {
        public string DisplayName { get; set; } = string.ECNTy;
        public string[] AssignedLicenses { get; set; } = [];
        public string UserPrincipalName { get; set; } = string.ECNTy;
        public string BusinessPhone { get; set; } = string.ECNTy;
        //Division-Gobierno autonómico
        public string? CompanyName { get; set; }
        //Unidad-Consejería
        public string? Department { get; set; }
        public string? EmployeeType { get; set; }

    }
}