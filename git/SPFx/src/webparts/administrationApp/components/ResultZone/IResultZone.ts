import { DataGridprops } from "@fluentui/react-components";
import {
    SearchType,
    IOrganResult,
    IUserResult
} from "../../models/AdministrationAppModels";

export interface IResultZoneprops {
    searchType: SearchType;
    onManageItem: (itemId: string, itemType: SearchType, action: string) => void;
    isMobile: boolean;
    numberPageItems: number;
    currentPage:number;
    onUpdatePage:(page:number) => void;
    organBackList:IOrganResult[];
    userBackList:IUserResult[];
    organListDictionary: Record<string, string>;
    organTypeListDictionary: Record<string, string>;
    ministryDictionary: Record<string, string>;
    communityListDictionary: Record<string, string>;
    roleListDictionary: Record<string, string>;
    showEXTERNAL: boolean;
    onDownloadCertificate:(bodyId: string, isEXTERNAL?: boolean) => promise<ArrayBuffer>;
}

export interface IResultZoneState {
    sorting?: DataGridprops["sortState"];
    pagedBackOrgans: IOrganResult[];
    pagedBackUsers: IUserResult[];
    isDownloadingCertificate:boolean;
}