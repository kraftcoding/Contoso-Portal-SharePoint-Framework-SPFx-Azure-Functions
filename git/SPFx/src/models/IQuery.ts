export default interface IQuery {
    viewFields: string[];
    filter: string;
    expand: string[];
    top?: number;
}

export interface IQueryOrder extends IQuery {
    orderBy: string;
    orderByAscending: boolean;
}