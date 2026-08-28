import IQuery from "./IQuery";

export class StestcentDocument {
    name: string;
    remoteItem: {
        webDavUrl: string;
    }
}

export enum RecentDocumentInternalNames {
    name = "name",
    remoteItem = "remoteItem",
    webDavUrl = "webDavUrl"
}

export class RecentDocument extends StestcentDocument {

    public static getAll(): IQuery {
        return {
            viewFields: [],
            expand: [],
            filter: '',
        };
    }

    public static mapSPToObject(props: StestcentDocument): RecentDocument {
        return {
            name: props.name,
            remoteItem: props.remoteItem,
        };
    }

}