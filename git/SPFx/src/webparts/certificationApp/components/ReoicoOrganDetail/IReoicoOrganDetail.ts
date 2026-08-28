import { Term } from "../../../../models/ITag";
import { IRelationBusinessArea } from "../../../BusinessAreaRelations/components/BusinessAreaRelationsModels";
import { IErrorsForm, ILookupEXTERNALField, IManagedItem, IOrganEXTERNALResult} from "../../models/CertificationAppModels";

export interface IEXTERNALOrganDetailprops {
    managedItem: IManagedItem | any;
    onManageItem: (itemId: string,  action: string, item?:IOrganEXTERNALResult) => void;
    organTypeFilterOptions: ILookupEXTERNALField[];
    ministryFilterOptions: ILookupEXTERNALField[];
    organFilterOptions: Term[];
    BusinessAreaOptions:Term[];
    StatusOptions:ILookupEXTERNALField[];
    isMobile?:boolean;
    organDictionary: Record<string, string>;
    organTypeDictionary: Record<string, string>;
    ministryDictionary: Record<string, string>;
    errorUpdating?:boolean;
    allOrganCodes:string[];
    allOrgansData:IOrganEXTERNALResult[];
    BusinessAreaMiniesterioRelations: IRelationBusinessArea[];
}

export interface IEXTERNALDepartmentDetailstate {
    selectedTab: string;
    editSelectedItem: IOrganEXTERNALResult | any;
    showConfirmation?:boolean;
    showLoadingConfirmation?:boolean;
    formErrors?:IErrorsForm;
}