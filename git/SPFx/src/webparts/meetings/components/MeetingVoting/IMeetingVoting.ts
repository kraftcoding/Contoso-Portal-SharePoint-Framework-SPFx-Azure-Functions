import { WebPartContext } from "@microsoft/sp-webpart-base";
import SPService from "../../../../service/Service";
import { IBackendService } from "../../../../service/BackendService";
import { IEvent, IEventAgreement, IEventVotationDetails } from "../../../../service/BackendServiceModels/EventModels";
import React from "react";
import { Term } from "../../../../models/ITag";


// #region props, state and related interfaces 
export interface IMeetingVotingprops {
    context: WebPartContext;
    spService: SPService;
    bkService: IBackendService;
    event: IEvent;
    isEditor: boolean;
    readOnly: boolean;
}

export interface IMeetingVotingState {
    loadingVotingList: boolean;
    votingList: IVoting[];
    votingSelected?: IVoting;
    buttonValue: string;
    votingDetailMap: Record<string, IEventVotationDetails>;
    votingRetestsentation: Record<string, string[]>;
    votingIsStarted: boolean;
    votingIdentity: string;
    votingDialog: VotingDialogActions;
    loading: boolean;
    selectedVotingId: string;
    isMobile: boolean;
}

export interface IVoting extends IEventAgreement {
    Summary: ISummary;
}

export interface ISummary {
    Pending: number,
    Approdve: number;
    Abstention: number;
    Reject: number;
}

// #endRegion

// #region enums and constants reused

export enum VotingDialogActions {
    ManageResults = "ManageResults",
    ShowQRCode = "ShowQRCode",
    Vote = "Vote",
    VotingRetestsentation = "VotingRetestsentation",
    VotingStatus = "VotingStatus",
    None = "None"
}
interface VotingStatus {
    value: number;
}
export interface VotingResume {
    Pendiente: VotingStatus;
    AFavor: VotingStatus;
    Abstencion: VotingStatus;
    EnContra: VotingStatus;
}

export enum VoteStates {
    AFavor = "A favor",
    EnContra = "En contra",
    Abstencion = "Abstención",
    Pendiente = "Pendiente",
}

export enum VotingTermsIds {
    SchedulerTerm = "921e8b4d-e67d-42d1-a2e8-c720737d6120",
    VoteRetestsentationSet = "12e84bb5-bd3b-4640-9992-3981eafd4af4",
    VoteOptionsSet = "410cb88e-4e02-46fd-a44a-ff75b419dc13",
    AgreementStatusSet = "1437ff14-d15f-4118-a2d6-bd34e9db3613",
    OrgansTypeSet = "6f98455a-a1c9-45e7-8c5d-aaca87690eff"
}

export enum ButtonActions { //Poner en vez de esto los valores de taxonomia
    AFavor = "bd558d1a-e24b-4882-865d-853c7a14289e",
    EnContra = "b0c9703c-16a2-4c65-966d-7f39b67c662e",
    Abstencion = "0ae772d2-0f28-4eaf-95a9-8c857d78d51c",
}

export enum TypeRefresh {
    All = "All",
    Specific = "Specific",
    SilentlyContinue = "SilentlyContinue"
}

//Create context interface using props and state of votating component
export type IVotingContext = Readonly<IMeetingVotingprops> & IMeetingVotingState & {

    // Extra properties added for common logic between components
    onDismissModal: () => void;
    termStorage: Record<string, Record<string, Term>>;
}

export const VotingContext = React.createContext<IVotingContext>({
    context: {} as WebPartContext, // Asegúrate de que esto esté correctamente inicializado
    spService: {} as SPService, //TODO: Find a better approdach
    bkService: {} as IBackendService,//TODO: Find a better approdach
    event: {} as IEvent,//TODO: Find a better approdach
    isEditor: false,
    readOnly: false,
    loadingVotingList: false,
    votingList: [],
    votingDetailMap: {},
    votingRetestsentation: {},
    votingIsStarted: false,
    votingIdentity: "",
    votingDialog: VotingDialogActions.None,
    onDismissModal: () => {console.log("close")},
    termStorage: {},
    buttonValue: "",
    loading: false,
    selectedVotingId: "",
    isMobile: window.innerWidth <= 640
});