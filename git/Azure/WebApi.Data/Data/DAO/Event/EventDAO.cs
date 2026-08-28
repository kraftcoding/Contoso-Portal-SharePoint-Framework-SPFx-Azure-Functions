using Contoso.Portal.Data.DAL.Helpers;
using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public class BaseEventDAO : BaseSPListItemDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.EventBase;

        public string DocumentSetDescription { get; set; } = string.ECNTy;
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Location { get; set; } = string.ECNTy;
        public string IdMeeting { get; set; } = string.ECNTy;
        public string LocationDetails { get; set; } = string.ECNTy;
        public FieldTaxonomyValue? MeetingType { get; set; }
        public FieldTaxonomyValue? EstadoMeeting { get; set; }
        public FieldTaxonomyValue? Department { get; set; }
        public FieldTaxonomyValue? TipoDepartment { get; set; }
        public FieldTaxonomyValue? OnlineTool { get; set; }
        public FieldUrlValue? UrlOnlineTool { get; set; }
        public string OutlookMeetingId { get; set; } = string.ECNTy;
        public string IdReunionOnlineMSTeams { get; set; } = string.ECNTy;
        public FieldUserValue? GrupoAsistentesAsociado { get; set; }
        public FieldUserValue[]? Asistentes { get; set; }
        public string VersionPublicada { get; set; } = string.ECNTy;
        public DateTime? LastPublishedDate { get; set; }
        public List<string>? Recordatorios { get; set; }

        public string GetUbicacionCompleta()
        {
            
            string result = string.ECNTy;
            if(!string.IsNullOrECNTy(Location))
                result = Location;

            if(!string.IsNullOrECNTy(LocationDetails))
                result += " - " + LocationDetails;

            return result;
            
        }

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            item[nameof(DocumentSetDescription)] = DocumentSetDescription;
            item[nameof(StartDate)] = StartDate;
            item[nameof(EndDate)] = EndDate;
            item[nameof(Location)] = Location;
            item[nameof(LocationDetails)] = LocationDetails;
            item[nameof(MeetingType)] = MeetingType;
            item[nameof(EstadoMeeting)] = EstadoMeeting;
            item[nameof(Department)] = Department;
            item[nameof(TipoDepartment)] = TipoDepartment;
            item[nameof(OnlineTool)] = OnlineTool;
            item[nameof(UrlOnlineTool)] = UrlOnlineTool;
            item[nameof(OutlookMeetingId)] = OutlookMeetingId;
            item[nameof(IdReunionOnlineMSTeams)] = IdReunionOnlineMSTeams;
            item[nameof(GrupoAsistentesAsociado)] = SaveFixSPPrincipal.FixUserFieldValue(GrupoAsistentesAsociado);
            item[nameof(VersionPublicada)] = VersionPublicada;
            item[nameof(LastPublishedDate)] = LastPublishedDate;
            item[nameof(Recordatorios)] = Recordatorios;
            item[nameof(IdMeeting)] = IdMeeting;
            item[nameof(Asistentes)] = SaveFixSPPrincipal.FixUserFieldValue(Asistentes);
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            DocumentSetDescription = item[nameof(DocumentSetDescription)] as string ?? string.ECNTy;
            StartDate = item[nameof(StartDate)] is not null ? (DateTime)item[nameof(StartDate)] : throw new Exception($"{nameof(StartDate)} is null in item {item.Id}");
            EndDate = item[nameof(EndDate)] is not null ? (DateTime)item[nameof(EndDate)] : throw new Exception($"{nameof(EndDate)} is null in item {item.Id}");
            Location = item[nameof(Location)] as string ?? string.ECNTy;
            LocationDetails = item[nameof(LocationDetails)] as string ?? string.ECNTy;
            MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;
            EstadoMeeting = item[nameof(EstadoMeeting)] as FieldTaxonomyValue;
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;
            OnlineTool = item[nameof(OnlineTool)] as FieldTaxonomyValue;
            UrlOnlineTool = item[nameof(UrlOnlineTool)] as FieldUrlValue;
            OutlookMeetingId = item[nameof(OutlookMeetingId)] as string ?? string.ECNTy;
            IdReunionOnlineMSTeams = item[nameof(IdReunionOnlineMSTeams)] as string ?? string.ECNTy;
            GrupoAsistentesAsociado = item[nameof(GrupoAsistentesAsociado)] as FieldUserValue;
            VersionPublicada = item[nameof(VersionPublicada)] as string ?? string.ECNTy;
            LastPublishedDate = item[nameof(LastPublishedDate)] as DateTime?;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            Asistentes = (item[nameof(Asistentes)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
            Recordatorios = item[nameof(Recordatorios)] as List<string> ?? new List<string>();
        }

        public override void MapFromSPListItemVersion(IListItemVersion item)
        {
            base.MapFromSPListItemVersion(item);
            DocumentSetDescription = item[nameof(DocumentSetDescription)] as string ?? string.ECNTy;
            StartDate = item[nameof(StartDate)] is not null ? (DateTime)item[nameof(StartDate)] : throw new Exception($"{nameof(StartDate)} is null in item {item.Id}");
            EndDate = item[nameof(EndDate)] is not null ? (DateTime)item[nameof(EndDate)] : throw new Exception($"{nameof(EndDate)} is null in item {item.Id}");
            Location = item[nameof(Location)] as string ?? string.ECNTy;
            LocationDetails = item[nameof(LocationDetails)] as string ?? string.ECNTy;
            MeetingType = item[nameof(MeetingType)] as FieldTaxonomyValue;
            EstadoMeeting = item[nameof(EstadoMeeting)] as FieldTaxonomyValue;
            Department = item[nameof(Department)] as FieldTaxonomyValue;
            TipoDepartment = item[nameof(TipoDepartment)] as FieldTaxonomyValue;
            OnlineTool = item[nameof(OnlineTool)] as FieldTaxonomyValue;
            UrlOnlineTool = item[nameof(UrlOnlineTool)] as FieldUrlValue;
            OutlookMeetingId = item[nameof(OutlookMeetingId)] as string ?? string.ECNTy;
            IdReunionOnlineMSTeams = item[nameof(IdReunionOnlineMSTeams)] as string ?? string.ECNTy;
            GrupoAsistentesAsociado = item[nameof(GrupoAsistentesAsociado)] as FieldUserValue;
            VersionPublicada = item[nameof(VersionPublicada)] as string ?? string.ECNTy;
            LastPublishedDate = item[nameof(LastPublishedDate)] as DateTime?;
            IdMeeting = item[nameof(IdMeeting)] as string ?? string.ECNTy;
            Recordatorios = item[nameof(Recordatorios)] as List<string> ?? new List<string>();
            Asistentes = (item[nameof(Asistentes)] as IFieldValueCollection)?.Values.OfType<FieldUserValue>().ToArray();
        }

        public override Dictionary<string, object> AsNewListItem()
        {
            var values = base.AsNewListItem();
            values[nameof(DocumentSetDescription)] = DocumentSetDescription;
            values[nameof(StartDate)] = StartDate;
            values[nameof(EndDate)] = EndDate;
            values[nameof(Location)] = Location;
            values[nameof(LocationDetails)] = LocationDetails;
            values[nameof(MeetingType)] = MeetingType;
            values[nameof(EstadoMeeting)] = EstadoMeeting;
            values[nameof(OnlineTool)] = OnlineTool;
            values[nameof(UrlOnlineTool)] = UrlOnlineTool;
            values[nameof(OutlookMeetingId)] = OutlookMeetingId;
            values[nameof(IdReunionOnlineMSTeams)] = IdReunionOnlineMSTeams;
            values[nameof(GrupoAsistentesAsociado)] = SaveFixSPPrincipal.FixUserFieldValue(GrupoAsistentesAsociado);
            values[nameof(VersionPublicada)] = VersionPublicada;
            values[nameof(LastPublishedDate)] = LastPublishedDate;
            values[nameof(IdMeeting)] = IdMeeting;
            values[nameof(Recordatorios)] = Recordatorios;
            values[nameof(Asistentes)] = SaveFixSPPrincipal.FixUserFieldValue(Asistentes);

            return values;
        }

        public T AsChildCT<T>() where T : BaseEventDAO, new()
        {
            var child = new T();
            if (this.HasSPListItem())
            {
                child.MapFromSPListItem(this.AsListItem());
            }
            else
            {
                // when created as new or from version
                child.ID = this.ID;
                child.Title = this.Title;
                child.Author = this.Author;
                child.Editor = this.Editor;
                child.Created = this.Created;
                child.Modified = this.Modified;
                child.ContentTypeId = this.ContentTypeId;
                child.Location = this.Location;
                child.Department = this.Department;
                child.TipoDepartment = this.TipoDepartment;
                child.LocationDetails = this.LocationDetails;
                child.DocumentSetDescription = this.DocumentSetDescription;
                child.EndDate = this.EndDate;
                child.StartDate = this.StartDate;
                child.OutlookMeetingId = this.OutlookMeetingId;
                child.IdReunionOnlineMSTeams = this.IdReunionOnlineMSTeams;
                child.IdMeeting = this.IdMeeting;
                child.MeetingType = this.MeetingType;
                child.OnlineTool = this.OnlineTool;
                child.EstadoMeeting = this.EstadoMeeting;
                child.UrlOnlineTool = this.UrlOnlineTool;
                child.Asistentes = this.Asistentes;
                child.GrupoAsistentesAsociado = this.GrupoAsistentesAsociado;
                child.VersionPublicada = this.VersionPublicada;
                child.LastPublishedDate = this.LastPublishedDate;
                child.Recordatorios = this.Recordatorios;
            }
            return child;
        }
    }

    public class MeetingEnEdicionDAO : BaseEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.InConstructionEvent;

        public override Dictionary<string, object> AsNewListItem()
        {
            var values = base.AsNewListItem();
            values[nameof(IdMeeting)] = Guid.NewGuid().ToString("N"); // to ensure unique id across all libraries
            return values;
        }
    }

    public class MeetingPublicadaDAO : BaseEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.PublishedEvent;
    }

    public class MeetingArchivadaDAO : BaseEventDAO
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.ArchivedEvent;
    }
}