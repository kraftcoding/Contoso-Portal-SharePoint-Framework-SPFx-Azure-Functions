import { IQueryOrder } from "./IQuery";

export class SPUsefulLink {
    Title: string;
    URL: {
        Url: string;
    };
    NuevaVentana: boolean;
}

export enum UsefulLinkNames {
    Title = 'Title',
    URL = 'URL',
    NuevaVentana = 'NuevaVentana'
}

export class UsefulLink extends SPUsefulLink {

    public static readonly ListUrl: string = '/Lists/Enlaces';

    public static getAll(): IQueryOrder {
        return {
            viewFields: [
                UsefulLinkNames.Title,
                UsefulLinkNames.URL,
                UsefulLinkNames.NuevaVentana
            ],
            expand: [],
            filter: '',
            orderBy: UsefulLinkNames.Title,
            orderByAscending: true
        };

    }

    public static mapSPToObject(props: SPUsefulLink): UsefulLink {
        return {
            Title: props.Title,
            URL: props.URL,
            NuevaVentana: props.NuevaVentana
        };
    }

}