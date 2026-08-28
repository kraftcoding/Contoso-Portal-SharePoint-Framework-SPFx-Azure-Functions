import IQuery from "./IQuery";

export class SPTask {
    Title: string;
    date: Date;
    link: string;
}

export enum TaskInternalNames {
}

export class Task extends SPTask {
    public static getAll(): IQuery {
        return {
            viewFields: [],
            expand: [],
            filter: '',
        };
    }

    public static mapSPToObject(props: SPTask): Task {
        return {
            Title: props.Title,
            date: props.date,
            link: props.link
        };
    }
}