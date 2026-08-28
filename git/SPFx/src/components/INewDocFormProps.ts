import { WebPartContext } from "@microsoft/sp-webpart-base";
import SPService from "./../service/Service";

export interface INewDocFormprops {
    context: WebPartContext;
    spService: SPService;
    relativeUrlToGetDocs: string;
    open: boolean;
    relatedDocumentsIds: string[];
    onSaveData: (fileUniqueId: string) => promise<void>;
    openCloseForm: () => void;
    relatedItemTitle?: string;
}