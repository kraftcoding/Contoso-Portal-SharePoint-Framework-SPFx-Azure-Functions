namespace Contoso.Portal.Model.Bodies
{
    public class BusinessUnit
    {
        public int Id { get; set; } = 0;
        public string? BusinessArea { get; set; }
        public string? Division { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }

    }
}