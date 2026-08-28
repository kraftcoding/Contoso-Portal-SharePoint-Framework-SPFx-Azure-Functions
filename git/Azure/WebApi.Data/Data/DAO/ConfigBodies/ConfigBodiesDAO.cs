using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.ConfigDepartments
{
    public class ConfigDepartmentsDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.ConfiguracionDepartments;

        public FieldTaxonomyValue? Department { get; set; }

        public FieldTaxonomyValue? TipoDepartment { get; set; }

        public FieldTaxonomyValue? Secretaria { get; set; }

        public FieldTaxonomyValue? Division { get; set; }

        public string? dir { get; set; }

        public string? SIA { get; set; }

        public string? Materia { get; set; }

        public string? Abreviatura { get; set; }

        public FieldTaxonomyValue? Period { get; set; }

        public FieldTaxonomyValue? BusinessArea { get; set; }

        //public int? Departmentsuperior { get; set; }

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

        //public int IdCertificateBody { get; set; }

        public int IdAgendaBodies { get; set; }

        public int IdAttendanceBodies { get; set; }

        public int? DiasAprodbacionMinutes { get; set; }

        public string? DocumentSetDescription { get; set; }

        public bool Intersectorial { get; set; }

        public FieldTaxonomyValue? TipoMembresia { get; set; }

        public string? IdentificadorConferencia { get; set; }



        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(Department)] = Department;
            item[nameof(TipoDepartment)] = TipoDepartment;
            item[nameof(Secretaria)] = Secretaria;
            item[nameof(Division)] = Division;
            item[nameof(dir)] = dir;
            item[nameof(SIA)] = SIA;
            item[nameof(Materia)] = Materia;
            item[nameof(Abreviatura)] = Abreviatura;
            item[nameof(Period)] = Period;
            item[nameof(BusinessArea)] = BusinessArea;
            //item[nameof(Departmentsuperior)] = Departmentsuperior;
            item[nameof(Activo)] = Activo;
            item[nameof(FechaConstitucion)] = FechaConstitucion;
            item[nameof(FechaExtincion)] = FechaExtincion;
            item[nameof(Observaciones)] = Observaciones;
            item[nameof(IdEmailBodies)] = IdEmailBodies;
            item[nameof(IdEmailBodiesOnline)] = IdEmailBodiesOnline;
            item[nameof(IdEmailBodiesOnlineInPerson)] = IdEmailBodiesOnlineInPerson;
            item[nameof(IdEmailBodiesInPerson)] = IdEmailBodiesInPerson;
            item[nameof(IdEmailBodiesWrittenprodcedure)] = IdEmailBodiesWrittenprodcedure;
            item[nameof(IdEmailBodiesDocumentationReferral)] = IdEmailBodiesDocumentationReferral;
            item[nameof(IdActBodies)] = IdActBodies;
            item[nameof(IdAgendaBodies)] = IdAgendaBodies;
            item[nameof(IdAttendanceBodies)] = IdAttendanceBodies;
            item[nameof(IdCertificateBodies)] = IdCertificateBodies;
            //item[nameof(IdCertificateBody)] = IdCertificateBody;
            item[nameof(DiasAprodbacionMinutes)] = DiasAprodbacionMinutes;
            item[nameof(DocumentSetDescription)] = DocumentSetDescription;
            item[nameof(Intersectorial)] = Intersectorial;
            item[nameof(TipoMembresia)] = TipoMembresia;
            item[nameof(IdentificadorConferencia)] = IdentificadorConferencia;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;
            Secretaria = item[nameof(Secretaria)] as FieldTaxonomyValue;
            Division = item[nameof(Division)] as FieldTaxonomyValue;
            dir = item[nameof(dir)] as string ?? string.ECNTy;
            SIA = item[nameof(SIA)] as string ?? string.ECNTy;
            Materia = item[nameof(Materia)] as string ?? string.ECNTy;
            Abreviatura = item[nameof(Abreviatura)] as string ?? string.ECNTy;
            Period = item[nameof(Period)] as FieldTaxonomyValue;
            BusinessArea = item[nameof(BusinessArea)] as FieldTaxonomyValue;
            //Departmentsuperior = item[nameof(Departmentsuperior)] as int ?? 0;
            Activo = (bool)item[nameof(Activo)];
            FechaConstitucion = item[nameof(FechaConstitucion)] is not null ? (DateTime)item[nameof(FechaConstitucion)] : throw new Exception($"{nameof(FechaConstitucion)} is null in item {item.Id}");
            FechaExtincion = item[nameof(FechaExtincion)] is not null ? (DateTime)item[nameof(FechaExtincion)] : throw new Exception($"{nameof(FechaExtincion)} is null in item {item.Id}");
            Observaciones = item[nameof(Observaciones)] as string ?? string.ECNTy;
            IdEmailBodies = Convert.ToInt32(item[nameof(IdEmailBodies)]);
            IdEmailBodiesOnline = Convert.ToInt32(item[nameof(IdEmailBodiesOnline)]);
            IdEmailBodiesOnlineInPerson = Convert.ToInt32(item[nameof(IdEmailBodiesOnlineInPerson)]);
            IdEmailBodiesInPerson = Convert.ToInt32(item[nameof(IdEmailBodiesInPerson)]);
            IdEmailBodiesWrittenprodcedure = Convert.ToInt32(item[nameof(IdEmailBodiesWrittenprodcedure)]);
            IdEmailBodiesDocumentationReferral = Convert.ToInt32(item[nameof(IdEmailBodiesDocumentationReferral)]);
            IdActBodies = Convert.ToInt32(item[nameof(IdActBodies)]);
            IdAgendaBodies = Convert.ToInt32(item[nameof(IdAgendaBodies)]);
            IdCertificateBodies = Convert.ToInt32(item[nameof(IdCertificateBodies)]);
            IdAttendanceBodies = Convert.ToInt32(item[nameof(IdAttendanceBodies)]);
            //IdCertificateBody = Convert.ToInt32(item[nameof(IdCertificateBody)]);
            DiasAprodbacionMinutes = item[nameof(DiasAprodbacionMinutes)] is not null ? Convert.ToInt32(item[nameof(DiasAprodbacionMinutes)]) : 0;
            DocumentSetDescription = item[nameof(DocumentSetDescription)] as string ?? string.ECNTy;
            Intersectorial = (bool)item[nameof(Intersectorial)];
            TipoMembresia = item[nameof(TipoMembresia)] as FieldTaxonomyValue;
            IdentificadorConferencia = item[nameof(IdentificadorConferencia)] as string ?? string.ECNTy;
        }
    }
}