import { ItemReference } from "@microsoft/microsoft-graph-types";
export interface IDocumentsState {
    breadCrumbItems: IBreadCrumbItem[];
    currentDriveItemId?: IBreadCrumbItem;
}

export interface IBreadCrumbItem {
    title: string;
    id?: string;
    parent?: ItemReference;
}