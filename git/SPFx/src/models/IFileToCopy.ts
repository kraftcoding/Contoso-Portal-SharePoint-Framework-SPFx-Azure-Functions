import IQuery from "./IQuery";

export enum MeetingsInternalNames {
    Title = 'Title',
    FileRef = 'FileRef'
}

export interface IFileToCopy {
    Id: string;
    Title: string;
    Path: string;
    Type: string;
}

export interface IMeetings{
    Title: string;
    FileRef: string;
}

export class Meetings {
    public static ConvosListName: string = '/MeetingsConstruccion';
    
    public static getAllMeetings(): IQuery {
        return {
            viewFields: [MeetingsInternalNames.FileRef,MeetingsInternalNames.Title],
            expand: [],
            filter: ''
        };
    }
}