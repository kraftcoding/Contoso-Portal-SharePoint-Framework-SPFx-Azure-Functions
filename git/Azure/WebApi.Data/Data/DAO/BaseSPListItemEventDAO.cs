using PnP.Core.Model.SharePoint;

namespace Contoso.Portal.Data.DAO;

public abstract class BaseSPListItemEventDAO : BaseSPListItemDAO
{
    public FieldTaxonomyValue? Idioma { get; set; } //TODO: Include idioma

    public FieldTaxonomyValue? Department { get; set; }

    public FieldTaxonomyValue? TipoDepartment { get; set; }

    public string? IdMeeting { get; set; }
    public FieldTaxonomyValue? MeetingType { get; set; }

    public string? UniqueSharedID { get; set; }

    public void MapToSPListItemEvent(IListItem item)
    {
        item[nameof(UniqueSharedID)] = string.IsNullOrECNTy(UniqueSharedID) ? Guid.NewGuid().ToString("N") : UniqueSharedID;

        // Campos que se autorrellenan del DocumentSet
        //Idioma 
        //IdMeeting
        //Department
        //Tipo de Department
    }

    public void MapFromSPListItemEvent(IListItem item)
    {
        Idioma = item[nameof(Idioma)] as FieldTaxonomyValue;
        Department = item[nameof(Department)] as FieldTaxonomyValue;
        TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;

        IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
        MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;

        UniqueSharedID = item[nameof(UniqueSharedID)] as string ?? Guid.NewGuid().ToString("N");
    }

    public void MapFromSPListItemEventVersion(IListItemVersion item)
    {
        Idioma = item[nameof(Idioma)] as FieldTaxonomyValue;
        Department = item[nameof(Department)] as FieldTaxonomyValue;
        TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;

        IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
        MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;

        UniqueSharedID = item[nameof(UniqueSharedID)] as string ?? Guid.NewGuid().ToString("N");
    }

    public IDictionary<string, object> AsNewListItemEvent(IDictionary<string, object> values)
    {
        values[nameof(UniqueSharedID)] = string.IsNullOrECNTy(UniqueSharedID) ? Guid.NewGuid().ToString("N") : UniqueSharedID;
        return values;
    }
}