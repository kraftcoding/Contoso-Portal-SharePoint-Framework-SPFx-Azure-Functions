import { DataGridprops } from "@fluentui/react-components";
import {
    IOrganEXTERNALResult
} from "../../models/CertificationAppModels";

export interface IResultZoneprops {
    onManageItem: (itemId: string, action: string) => void;
    isMobile: boolean;
    numberPageItems: number;
    currentPage:number;
    onUpdatePage:(page:number) => void;
    organBackList:IOrganEXTERNALResult[];
    organListDictionary: Record<string, string>;
    organTypeListDictionary: Record<string, string>;
    ministryDictionary: Record<string, string>;
    onDownloadCertificate:(bodyId: string, isEXTERNAL?: boolean) => promise<ArrayBuffer>;
}

export interface IResultZoneState {
    sorting?: DataGridprops["sortState"];
    pagedBackOrgans: IOrganEXTERNALResult[];
    isDownloadingCertificate:boolean;
}