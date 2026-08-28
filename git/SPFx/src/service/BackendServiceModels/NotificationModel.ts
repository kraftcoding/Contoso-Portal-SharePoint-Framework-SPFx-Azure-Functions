import { IPendingChanges } from "../../models/IPendingChanges";

export interface INotification {
    Id: string;
    Title: string;
    Body: string;
    Expires: Date;
    Visualized: boolean;
    Source: string;
    Url: string;
    Locale: string;
    NotificationType: string;
    NotificationPriority: string;
    Created?: string;
    PendingChanges?: IPendingChanges[];
}

export interface INotificationBathRequest {
    Operation: NotificationBatchOp;
    NotificationIds: string[];
}

export enum NotificationBatchOp {
    Read = 0,
    Unread = 1,
    Hide = 2
}