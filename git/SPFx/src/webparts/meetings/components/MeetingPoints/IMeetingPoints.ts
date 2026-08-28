import { Term } from '../../../../models/ITag';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IEvent, IEventAgendaItem } from "../../../../service/BackendServiceModels/EventModels";
import SPService from "../../../../service/Service";
import { IBackendService } from "../../../../service/BackendService";
import { IFileInfo } from '@pnp/sp/files';

export interface IMeetingPointsprops {
    event: IEvent;
    typeGroupId: string;
    typeSetId: string;
    context: WebPartContext;
    spService: SPService;
    bkService: IBackendService;
    isEditor: boolean;
    readOnly: boolean;
}

export interface IMeetingPointsState {
    isLoading: boolean;
    isLoadingItemId?: string;
    agendaItems: IEventAgendaItem[];
    agendaItemsDictionary: { [key: string]: Array<IEventAgendaItem> };
    agendaItemTypes: Term[];
    agendaItemOrderTypes: Term[];
    editAgendaItem?: IEventAgendaItem;
    aNewAgendaItemIsBeingEdited?: boolean;
    openNewDoc: boolean;
    sectionNewDoc?: IEventAgendaItem;
    selectedCertificateAdd?:IEventAgendaItem;
    fileNames: IFileInfo[];
    errors: IAgendaItemErrors;
    validate: boolean;
    downloadingTheOrderOfTheDay: boolean;
    isInReOrderMode: boolean;
    multiDialogState: IMultiDialogState;
}

export interface IMultiDialogState{
    isOpen: boolean;
    type?: MultiDialogType;
    props: {
        document?: IFileInfo;
        agendaItem?: IEventAgendaItem;
        fileId?: string;
    };
}

export interface IAgendaItemErrors {
    title?: string;
    duration?: string;
}

export interface ITooltipConfig {
    [key: string]: {
        isDisabled: boolean;
        message: string;
    }
}
export enum MultiDialogType {
    CancelReorder = "CancelReorder",
    SaveReorder = "SaveReorder",
    RemoveDocument = "RemoveDocument",
    RemoveOrder = "RemoveOrder"
}
export enum ItemActions {
    LvlUp = "LvlUp",
    LvlDown = "LvlDown",
    ArrowUp = "ArrowUp",
    ArrowDown = "ArrowDown",
    Remove = "Remove",
    Add = "Add",
    Edit = "Edit",
    Attach = "Attach",
    UploadCertificate = "UploadCertificate",
    DownloadCertificate = "DownloadCertificate"
}

export enum LineTypes {
    I = "I",
    T = "T",
    S = "S",
    L = "L"
}