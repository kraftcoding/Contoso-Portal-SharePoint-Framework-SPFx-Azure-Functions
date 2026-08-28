namespace Contoso.Portal.Model.Bodies
{
    public class ConfigDepartments
    {
        public string? BodyId { get; set; }
        public string? NombreDepartment { get; set; }
        public string? Department { get; set; }
        public string? TipoDepartment { get; set; }
        public string? Secretaria { get; set; }
        public string? Division { get; set; }
        public string? dir { get; set; }
        public string? SIA { get; set; }
        public string? Materia { get; set; }
        public string? Abreviatura { get; set; }
        public string? Period { get; set; }
        public string? BusinessArea { get; set; }
        //public int Departmentsuperior { get; set; }
        public bool Activo { get; set; }
        public DateTime? FechaConstitucion { get; set; }
        public DateTime? FechaExtincion { get; set; }
        public string? Observaciones { get; set; }
        public int IdEmailBodies { get; set; }
        public int IdEmailBodiesOnline { get; set; }
        public int IdEmailBodiesOnlineInPerson { get; set; }
        public int IdEmailBodiesInPerson { get; set; }
        public int IdEmailBodiesWrittenprodcedure { get; set; }
        public int IdEmailBodiesDocumentationReferral { get; set; }
        public int IdActBodies { get; set; }
        public int IdCertificateBodies { get; set; }
        public int IdCertificateBody { get; set; }
        public int IdAgendaBodies { get; set; }
        public int IdAttendanceBodies { get; set; }
        public int? DiasAprodbacionMinutes { get; set; }
        public List<User>? Users { get; set; }
        public string? Description { get; set; }
        public bool Intersectorial { get; set; }
        public string? TipoMembresia { get; set; }
        public string? IdentificadorConferencia { get; set; }
        public bool? IsEXTERNAL { get; set; }

    }
}