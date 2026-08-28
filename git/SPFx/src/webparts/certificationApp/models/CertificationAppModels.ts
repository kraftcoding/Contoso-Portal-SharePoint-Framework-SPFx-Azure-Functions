import { IQueryOrder } from "../../../models/IQuery";

export interface IOrganEXTERNALResult {
    CodigoDepartment: string;
    BodyId: string;
    ID?:string;
    Department: string;
    TipoDepartmentEXTERNAL: string;
    SecretariaText: string;
    Division?: string;
    DocumentSetDescription:string;
    Intersectorial:boolean;
    dir: string;
    SIA: string;
    Abreviatura: string;
    Activo: boolean;
    FechaConstitucion?: string;
    FechaExtincion?: string;
    FechaInscripcion?:string;
    Observaciones: string;
    IdCertificateBodies: number;
    BusinessArea: string;
    NombreDepartment:string;
    StatusDepartment:string;
    FileRef?: string;
    Inscrito:boolean;
    FechaCreacion?: string;
    DepartmentAdscripcion?: string;
}


export interface IManagedItem {
    itemId: string;
    organDialogContent: IOrganEXTERNALResult | undefined;
    isNewItem?:boolean;
}

export interface IObjectDifferences {
    property: string;
    oldValue: string;
    newValue: string;
}

export interface IErrorsForm {
    CodigoDepartment?: string;
    BodyId?: string;
    Department?: string;
    TipoDepartmentEXTERNAL?: string;
    SecretariaText?: string;
    Division?: string;
    DocumentSetDescription?:string;
    Intersectorial?:boolean;
    dir?: string;
    SIA?: string;
    Abreviatura?: string;
    Activo?: boolean;
    FechaConstitucion?: string;
    FechaExtincion?: string;
    FechaInscripcion?:string;
    Observaciones?: string;
    IdCertificateBodies?: number;
    BusinessArea?: string;
    NombreDepartment?:string;
    StatusDepartment?:string;
}

export enum OrganTaxonomyIds{
    OrganType = 'ba3ff2ef-09c8-4345-9bad-777014ddf636',
    Organ = '6f98455a-a1c9-45e7-8c5d-aaca87690eff',
    Ministry = '4581a675-bfd7-4297-a403-5a7cce1be8e6',
    Secretaria = 'c5df120a-8954-43d7-9f71-fdd0e996ee30',
    Period = '7b4c61b9-c43e-4d38-84f6-f717b2bddcac',
    Community = '172d84e2-ba88-40d6-a872-8fe00d066caa',
    Roles = 'c107d212-3417-472b-a979-b9da999e6adf',
    BusinessArea = 'bda21274-3507-4b59-89e4-64cb5bb1aa0c',
    OrganTypeEXTERNAL = '42843e95-a65d-4db4-98f1-cf7571aaeb42',
    StatusDepartment = 'f0b6cf80-5014-42d7-8b35-73c132ca246c',
    NullTaxonomy = '00000000-0000-0000-0000-000000000000'
}

export enum SearchFilterKeys {
    OrganTypeEXTERNAL = 'OrganTypeEXTERNAL',
    BusinessArea = 'BusinessArea',
    CodigoDepartment= 'CodigoDepartment',
    FechaInscripcionHasta = 'FechaInscripcionHasta',
    FechaConstitucionHasta = 'FechaConstitucionHasta',
    FechaExtincionHasta = 'FechaExtincionHasta',
    FechaInscripcionDesde = 'FechaInscripcionDesde',
    FechaConstitucionDesde = 'FechaConstitucionDesde',
    FechaExtincionDesde= 'FechaExtincionDesde'
}

export interface ISearchFilters {
    [SearchFilterKeys.OrganTypeEXTERNAL]: string[];
    [SearchFilterKeys.BusinessArea]: string[];
    [SearchFilterKeys.CodigoDepartment]: string;
    [SearchFilterKeys.FechaInscripcionHasta]: string;
    [SearchFilterKeys.FechaConstitucionHasta]: string;
    [SearchFilterKeys.FechaExtincionHasta]: string;
    [SearchFilterKeys.FechaInscripcionDesde]: string;
    [SearchFilterKeys.FechaConstitucionDesde]: string;
    [SearchFilterKeys.FechaExtincionDesde]: string;
}

export interface ILookupEXTERNALField {
    Title: string;
    Id: string;
    descripcion: string;
    activo: string;     
}
export class EXTERNALList {
public static readonly ListDivisionUrl: string = '/Lists/Division';
public static readonly ListPeriodUrl: string = '/Lists/Period';
public static readonly ListTipoDepartmentUrl: string = '/Lists/TipoDepartment';
public static readonly ListStatusUrl: string = '/Lists/Status';
public static readonly ListTipoDocumentUrl: string = '/Lists/TipoDocument';

    public static getAll(): IQueryOrder {
        return {
            viewFields: [
                "Title",
                "descripcion",
                "Id",
                "activo"
            ],
            expand: [],
            filter: '',
            orderBy: "Title",
            orderByAscending: true
        };

    }
}

export class Certification {
    public static OrganListName: string = '/DepartmentsEXTERNAL';

    public static eCNTyFilters: ISearchFilters = {
        [SearchFilterKeys.OrganTypeEXTERNAL]: [],
        [SearchFilterKeys.BusinessArea]: [],
        [SearchFilterKeys.CodigoDepartment]: "",
        [SearchFilterKeys.FechaInscripcionHasta]: "",
        [SearchFilterKeys.FechaConstitucionHasta]: "",
        [SearchFilterKeys.FechaExtincionHasta]: "",
        [SearchFilterKeys.FechaInscripcionDesde]: "",
        [SearchFilterKeys.FechaConstitucionDesde]: "",
        [SearchFilterKeys.FechaExtincionDesde]: ""
    };
}