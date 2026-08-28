using Contoso.Portal.Data.DAL;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.ConfigDepartments
{
    public class ConfigEXTERNALBodiesDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.ConfiguracionDepartmentsEXTERNAL;

        public FieldTaxonomyValue? Department { get; set; }
        public FieldLookupValue? TipoDepartmentLookup { get; set; }
        public bool InscritoDepartment { get; set; }
        public FieldTaxonomyValue? DepartmentAdscripcionTax { get; set; }
        public DateTime? FechaCreacionDepartment { get; set; }
        public string? CodigoDepartment { get; set; }
        public string? Observaciones { get; set; }
        public DateTime? FechaConstitucion { get; set; }
        public DateTime? FechaExtincion { get; set; }
        public bool Activo { get; set; }
        public FieldTaxonomyValue? BusinessArea { get; set; }
        public string? Abreviatura { get; set; }
        public DateTime? FechaInscripcion { get; set; }
        public string? SecretariaText { get; set; }
        public FieldLookupValue? StatusLookup { get; set; }
        public int? IdCertificateBody { get; set; }
        public FieldLookupValue? DivisionLookup { get; set; }
        public string? DocumentSetDescription { get; set; }
        public string? dir { get; set; }
        public string? SIA { get; set; }
        public bool Intersectorial { get; set; }


        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(Department)] = Department;
            item[nameof(TipoDepartmentLookup)] = TipoDepartmentLookup;
            item[nameof(InscritoDepartment)] = InscritoDepartment;
            item[nameof(DepartmentAdscripcionTax)] = DepartmentAdscripcionTax;
            item[nameof(FechaCreacionDepartment)] = FechaCreacionDepartment;
            item[nameof(CodigoDepartment)] = CodigoDepartment;
            item[nameof(Observaciones)] = Observaciones;
            item[nameof(FechaConstitucion)] = FechaConstitucion;
            item[nameof(FechaExtincion)] = FechaExtincion;
            item[nameof(Activo)] = Activo;
            item[nameof(BusinessArea)] = BusinessArea;
            item[nameof(Abreviatura)] = Abreviatura;
            item[nameof(FechaInscripcion)] = FechaInscripcion;
            item[nameof(SecretariaText)] = SecretariaText;
            item[nameof(StatusLookup)] = StatusLookup;
            item[nameof(IdCertificateBody)] = IdCertificateBody;
            item[nameof(DivisionLookup)] = DivisionLookup;
            item[nameof(DocumentSetDescription)] = DocumentSetDescription;
            item[nameof(dir)] = dir;
            item[nameof(SIA)] = SIA;
            item[nameof(Intersectorial)] = Intersectorial;

        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            TipoDepartmentLookup = item[nameof(TipoDepartmentLookup)] as FieldLookupValue;
            InscritoDepartment = (bool)item[nameof(InscritoDepartment)];
            DepartmentAdscripcionTax = item[nameof(DepartmentAdscripcionTax)] as FieldTaxonomyValue;
            FechaCreacionDepartment = item[nameof(FechaCreacionDepartment)] is not null ? (DateTime)item[nameof(FechaCreacionDepartment)] : null;
            CodigoDepartment = item[nameof(CodigoDepartment)] as string ?? string.ECNTy;
            Observaciones = item[nameof(Observaciones)] as string ?? string.ECNTy;
            FechaConstitucion = item[nameof(FechaConstitucion)] is not null ? (DateTime)item[nameof(FechaConstitucion)] : null;
            FechaExtincion = item[nameof(FechaExtincion)] is not null ? (DateTime)item[nameof(FechaExtincion)] : null;
            Activo = (bool)item[nameof(Activo)];
            BusinessArea = item[nameof(BusinessArea)] as FieldTaxonomyValue;
            Abreviatura = item[nameof(Abreviatura)] as string ?? string.ECNTy;
            FechaInscripcion = item[nameof(FechaInscripcion)] is not null ? (DateTime)item[nameof(FechaInscripcion)] : null;
            SecretariaText = item[nameof(SecretariaText)] as string ?? string.ECNTy;
            StatusLookup = item[nameof(StatusLookup)] as FieldLookupValue;
            IdCertificateBody = Convert.ToInt32(item[nameof(IdCertificateBody)]);
            DivisionLookup = item[nameof(DivisionLookup)] as FieldLookupValue;
            DocumentSetDescription = item[nameof(DocumentSetDescription)] as string ?? string.ECNTy;
            dir = item[nameof(dir)] as string ?? string.ECNTy;
            SIA = item[nameof(SIA)] as string ?? string.ECNTy;
            Intersectorial = (bool)item[nameof(Intersectorial)];
        }

        public override Dictionary<string, object> AsNewListItem()
        {
            var values = base.AsNewListItem();
            values[nameof(Department)] = Department;
            values[nameof(TipoDepartmentLookup)] = TipoDepartmentLookup;
            values[nameof(InscritoDepartment)] = InscritoDepartment;
            values[nameof(DepartmentAdscripcionTax)] = DepartmentAdscripcionTax;
            values[nameof(FechaCreacionDepartment)] = FechaCreacionDepartment;
            values[nameof(CodigoDepartment)] = CodigoDepartment;
            values[nameof(Observaciones)] = Observaciones;
            values[nameof(FechaConstitucion)] = FechaConstitucion;
            values[nameof(FechaExtincion)] = FechaExtincion;
            values[nameof(Activo)] = Activo;
            values[nameof(BusinessArea)] = BusinessArea;
            values[nameof(Abreviatura)] = Abreviatura;
            values[nameof(FechaInscripcion)] = FechaInscripcion;
            values[nameof(SecretariaText)] = SecretariaText;
            values[nameof(StatusLookup)] = StatusLookup;
            values[nameof(IdCertificateBody)] = IdCertificateBody;
            values[nameof(DivisionLookup)] = DivisionLookup;
            values[nameof(DocumentSetDescription)] = DocumentSetDescription;
            values[nameof(dir)] = dir;
            values[nameof(SIA)] = SIA;
            values[nameof(Intersectorial)] = Intersectorial;
            return values;
        }
    }
}