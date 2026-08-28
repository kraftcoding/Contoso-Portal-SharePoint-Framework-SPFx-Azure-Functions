
export enum RelationAreaFields{
    Id = 'Id',
    BusinessArea = 'BusinessArea',
    Division = 'Division',
    StartDate = 'StartDate',
    EndDate = 'EndDate',
    HasError = 'HasError'
}

export interface IRelationBusinessArea{
    [RelationAreaFields.Id]?:string;
    [RelationAreaFields.BusinessArea]: string;
    [RelationAreaFields.Division]: string;
    [RelationAreaFields.StartDate]: string;
    [RelationAreaFields.EndDate]: string;
    [RelationAreaFields.HasError]?: boolean;
}

export interface ILookupMinistryField {
    Title: string;
    Id: string;
    descripcion: string;
    activo: string;     
}

export enum RelationsTaxonomyIds{
    Division = '4581a675-bfd7-4297-a403-5a7cce1be8e6',
    BusinessArea = 'bda21274-3507-4b59-89e4-64cb5bb1aa0c',
}

export enum RelationsFilterKeys {
    BusinessArea = 'BusinessArea',
    Division = 'Division',
    Actual = 'Actual'
}

export interface IRelationsFilters {
    [RelationsFilterKeys.BusinessArea]: string[];
    [RelationsFilterKeys.Division]: string[];
    [RelationsFilterKeys.Actual]: boolean;
}

export class RelationsAreaMinistry {
    public static eCNTyFilters: IRelationsFilters = {
        [RelationsFilterKeys.BusinessArea]: [],
        [RelationsFilterKeys.Division]: [],
        [RelationsFilterKeys.Actual]: true
    };
}