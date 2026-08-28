import SPService from "../../../../service/Service";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IBackendService } from "../../../../service/BackendService";
import { IEvent, IEventStorage } from "../../../../service/BackendServiceModels/EventModels";
import { ItemReference, DriveItem } from "@microsoft/microsoft-graph-types";

export interface IMeetingDocsprops {
    context: WebPartContext;
    spService: SPService;
    bkService: IBackendService;
    event: IEvent;
    isEditor: boolean;
    readOnly: boolean;
}

export interface IMeetingDocsState {
    breadCrumbItems: IBreadCrumbItem[];
    currentDriveItemId?: IBreadCrumbItem;
    loadingDocuments: boolean;
    openNewDocForm: boolean;
    eventRootStorage?: IEventStorage;
    selectedFile?: DriveItem;
    deleting: boolean;
    allDocs: IMeetingDoc[];
    driveDocs: DriveItem[];
    backDocs: IDocumentSensitivityLabel[];
    sensitibityLabels:ITenantSensitivityLabel[];
    fileToUpdate?:IMeetingDoc;
    updatingDoc:boolean;
}

export interface IBreadCrumbItem {
    title: string;
    id?: string;
    parent?: ItemReference;
}

export interface IMeetingDoc extends DriveItem {
    //Id: string;
    //Filename: string;
    //AgendaItem: number;
    sensitivityLabelId?:string;
    sensitivityLabel?:string;
    isSignedDocument?:boolean;
    isExternalLabel?:boolean;
    isprocessingLabel?:boolean;
    isOtherError :boolean;
}

export enum IDocLabelColors{
    NoLabel ='grey',
    Applied = 'green',
    processing = 'orange',
    External = '#e24040'
}

export interface ITenantSensitivityLabel{
    Id: string;
    Name: string;
}

export interface IDocumentSensitivityLabel
{
    BodyId :string;
    SharedEventId :string;
    FilePath :string;
    SensitivityLabelId :string;
    IsExternalLabel :boolean;
    IsSignedDocument :boolean;
    IsOtherError :boolean;
    IsprocessingLabel :boolean;
    ExpectedLabel :string;
}