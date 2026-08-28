import { IFolderInfo } from "@pnp/sp/folders/types";
import { ISPService } from "../../service/Service";
import { IOrganEXTERNALResult } from "../../webparts/certificationApp/models/CertificationAppModels";

export interface IEXTERNALDocumentFormprops {
    isOpen: boolean;
    onClose: () => void;
    organs: IOrganEXTERNALResult[];
    organTypes: Record<string, string>;
    documentTypes: Record<string, string>;
    spService: ISPService;
}

export interface IEXTERNALDocumentFormState {
    document: IEXTERNALDoc;
    file?: File;
    isLoading: boolean;
    Departmentseleccionado?:IOrganEXTERNALResult;
    filteredOrgans:IOrganEXTERNALResult[];
    folders:IFolderInfo[];
    filterOrganTxt:string;
    selectedFolder?:IFolderInfo;
    errors: Record<string, string>;
    saved: boolean;
    hasError: boolean;
}

export interface IEXTERNALDoc{
    TipoDocumentLookup: string;
    TipoDepartmentLookup: string;
    Department: string;
    FechaMeeting: string;
    CarpetaReunion?: string;
    NuevaCarpeta?: boolean;
    NombreNuevaCarpeta?:string;
}

export interface IEXTERNALFolder{
    Name: string;
    Department: string;
    Url: string;
}