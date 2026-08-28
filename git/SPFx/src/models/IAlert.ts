import IQuery from "./IQuery";

export enum TipoAlert {
    Sistema = 'Sistema',
    Meetings = 'Meetings'
}

export class SPAlert {
    Title: string;
    Expires: Date | string;
    Body: string;
    TipoAlert: TipoAlert;
}

export enum AlertInternalNames {
    Title = 'Title',
    Expires = 'Expires',
    Body = 'Body',
    TipoAlert = 'TipoAlert'
}

export class Alert extends SPAlert {
    public static readonly ListUrl: string = '/Lists/Alerts';

    public static getAll(): IQuery {
        return {
            viewFields: [
                AlertInternalNames.Title,
                AlertInternalNames.Expires,
                AlertInternalNames.Body,
                AlertInternalNames.TipoAlert
            ],
            expand: [],
            filter: '',
        };
    }

    public static mapSPToObject(props: SPAlert): Alert {
        return {
            Title: props.Title,
            Expires: new Date(props.Expires as string),
            Body: props.Body,
            TipoAlert: props.TipoAlert
        };
    }
}