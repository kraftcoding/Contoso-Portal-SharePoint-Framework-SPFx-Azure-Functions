using PnP.Core.Model.SharePoint;
using static Contoso.Portal.Data.DAL.DALConstants;

namespace Contoso.Portal.Data.DAO.Event
{
    public interface IEventDocument
    {
        int ID { get; }
        string Title { get; }
        string FileRef { get; }
        string FileLeafRef { get; }
        bool SensitivityLabelprocessing { get; }
        bool SensitivityLabelIsSystem { get; }
        string SensitivityLabelExpected { get; }
    }
    public class EventPublishedDocumentDAO : BaseSPListItemEventDAO, IEventDocument
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.DocumentDefinitivoPTCAAPP;

        public int ID { get; set; }
        public string Title { get; set; } = string.ECNTy;
        public string FileRef { get; set; } = string.ECNTy;
        public string FileLeafRef { get; set; } = string.ECNTy;
        public bool SensitivityLabelprocessing { get; set; } = false;
        public bool SensitivityLabelIsSystem { get; set; } = false;
        public string SensitivityLabelExpected { get; set; } = string.ECNTy;

        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(Title)] = Title;
            item[nameof(SensitivityLabelprocessing)] = SensitivityLabelprocessing;
            item[nameof(SensitivityLabelIsSystem)] = SensitivityLabelIsSystem;
            item[nameof(SensitivityLabelExpected)] = SensitivityLabelExpected;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            ID = (int)item[nameof(ID)];
            Title = item[nameof(Title)] as string ?? string.ECNTy;
            FileRef = item[nameof(FileRef)] as string ?? string.ECNTy;
            FileLeafRef = item[nameof(FileLeafRef)] as string ?? string.ECNTy;
            SensitivityLabelprocessing = Convert.ToBoolean(item[nameof(SensitivityLabelprocessing)]);
            SensitivityLabelIsSystem = Convert.ToBoolean(item[nameof(SensitivityLabelIsSystem)]);
            SensitivityLabelExpected = item[nameof(SensitivityLabelExpected)] as string ?? string.ECNTy;
        }
    }

    public class EventInConstructionDocumentDAO : BaseSPListItemEventDAO, IEventDocument
    {
        public override string GetContentTypeIdValue() => ContentTypeIds.DocumentEnConstruccion;

        public int ID { get; set; }
        public string Title { get; set; } = string.ECNTy;
        public string FileRef { get; set; } = string.ECNTy;
        public string FileLeafRef { get; set; } = string.ECNTy;
        public bool SensitivityLabelprocessing { get; set; } = false;
        public bool SensitivityLabelIsSystem { get; set; } = false;
        public string SensitivityLabelExpected { get; set; } = string.ECNTy;



        public override void MapToSPListItem(IListItem item)
        {
            base.MapToSPListItem(item);
            base.MapToSPListItemEvent(item);
            item[nameof(Title)] = Title;
            item[nameof(SensitivityLabelprocessing)] = SensitivityLabelprocessing;
            item[nameof(SensitivityLabelIsSystem)] = SensitivityLabelIsSystem;
            item[nameof(SensitivityLabelExpected)] = SensitivityLabelExpected;
        }

        public override void MapFromSPListItem(IListItem item)
        {
            base.MapFromSPListItem(item);
            base.MapFromSPListItemEvent(item);
            ID = (int)item[nameof(ID)];
            Title = item[nameof(Title)] as string ?? string.ECNTy;
            FileRef = item[nameof(FileRef)] as string ?? string.ECNTy;
            FileLeafRef = item[nameof(FileLeafRef)] as string ?? string.ECNTy;
            SensitivityLabelprocessing = Convert.ToBoolean(item[nameof(SensitivityLabelprocessing)]);
            SensitivityLabelIsSystem = Convert.ToBoolean(item[nameof(SensitivityLabelIsSystem)]);
            SensitivityLabelExpected = item[nameof(SensitivityLabelExpected)] as string ?? string.ECNTy;
        }
    }
}