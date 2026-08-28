export interface IUser {
    loginName: string;
    isGuest: boolean;
}

export class Call {
    title?: string;
    startDate?: Date;
    endDate?: Date;
    location?: string;
    locationDescription?: string;
    online?: boolean;
    onlineTool?: string;
    organ?: string;
    urlTool?: string;
    users?: IUser[];
    description?: string;
}