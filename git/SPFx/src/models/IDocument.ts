import { IQueryOrder } from "./IQuery";

export class SPDocument {
    FileLeafRef: string;
    FileRef: string;
}

export enum DocumentInternalNames {
    FileLeafRef = "FileLeafRef",
    FileRef = "FileRef"
}

export class Document extends SPDocument {

    public static readonly ListUrl: string = '/Documents%20compartidos';

    public static getAll(): IQueryOrder {
        return {
            viewFields: [
                DocumentInternalNames.FileLeafRef,
                DocumentInternalNames.FileRef
            ],
            expand: [],
            filter: 'FSObjType eq 0',
            orderBy: DocumentInternalNames.FileLeafRef,
            orderByAscending: true
        };
    }

    public static mapSPToObject(props: SPDocument): Document {
        return {
            FileLeafRef: props.FileLeafRef,
            FileRef: props.FileRef
        };
    }

}