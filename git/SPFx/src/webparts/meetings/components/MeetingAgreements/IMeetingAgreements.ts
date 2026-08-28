import { WebPartContext } from "@microsoft/sp-webpart-base";
import SPService from "../../../../service/Service";
import { IBackendService } from "../../../../service/BackendService";
import { IEvent, IEventAgreement } from "../../../../service/BackendServiceModels/EventModels";
import { Term } from "../../../../models/ITag";
import { IFileInfo } from "@pnp/sp/files";

export interface IMeetingAgreementsprops {
    context: WebPartContext;
    spService: SPService;
    bkService: IBackendService;
    event: IEvent;
    isEditor: boolean;
    isGuest: boolean;
    readOnly: boolean;
    isOCprodle: boolean;
}

export interface IMeetingAgreementsState {
    isLoading: boolean;
    isLoadingItemId?: string;
    agreements: IEventAgreement[];
    pendingCertificateAgreementsIds: string[];
    fileNames: IFileInfo[];
    agreementStatuses: Term[];
    checkedAgreement?: IEventAgreement;
    selectedAgreementEditMode?: IEventAgreement;
    selectedAgreementAttachmentAdd?: IEventAgreement;
    selectedAgreementCertificateAdd?: IEventAgreement;
    openDocDialog: boolean;
    downloadingTemplate: boolean;
    selectedAgreementToRequestCertificate?: IEventAgreement;
    certificateRequestDialogIsOpen: boolean;
    sendingTheCertificateRequest: boolean;
    certificateRequestFailed: boolean;
    errorMessageInTheCertificateRequest: string;
}