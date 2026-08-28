import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingAgreements.module.scss';
import { IMeetingAgreementsprops, IMeetingAgreementsState } from './IMeetingAgreements';
import { ArrowDownload16Regular, ArrowDownload24Regular, DeleteRegular, Dismiss16Regular, DocumentRibbonRegular, EditRegular } from '@fluentui/react-icons';
import { Badge, Button, Checkbox, Field, Input, Radio, Spinner, Textarea, Tooltip } from '@fluentui/react-components';
import { Term } from '../../../../models/ITag';
import { ITermInfo } from '@pnp/sp/taxonomy/types';
import { getTermLabel } from '../../../../utils/Utils';
import { AgreementStatus, EventStatus, IEventAgreement } from '../../../../service/BackendServiceModels/EventModels';
import NewDocForm from '../../../../components/NewDocForm';
import { Logger } from '../../../../utils/Logger';
import { find, isECNTy } from '@microsoft/sp-lodash-subset';
import { File } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { IFileInfo } from '@pnp/sp/files';
import { ITaskCertification } from '../../../../service/BackendServiceModels/UserTaskModel';
import ConfirmAction from '../../../../components/ConfirmAction';

const AgreementStatesTermSetId = "1437ff14-d15f-4118-a2d6-bd34e9db3613";
AgreementStatus.Approdved
const ApprodvedStatusId: string = AgreementStatus.Approdved;
const ApprodvedUnanimouslyStatusId: string = AgreementStatus.UnanimouslyApprodved;
const DeclinedStatusId = AgreementStatus.Declined;
const OverTheTableStatusId = AgreementStatus.OverTheTable;

const InprodgressStatusId: string = AgreementStatus.Inprodgress;
const PendingStatusId: string = "00000000-0000-0000-0000-000000000000";

export default class MeetingAgreements extends React.Component<IMeetingAgreementsprops, IMeetingAgreementsState> {

    constructor(props: IMeetingAgreementsprops) {
        super(props);
        this.state = {
            isLoading: false,
            isLoadingItemId: undefined,
            agreements: [],
            pendingCertificateAgreementsIds: [],
            agreementStatuses: [],
            checkedAgreement: undefined,
            selectedAgreementEditMode: undefined,
            selectedAgreementAttachmentAdd: undefined,
            selectedAgreementCertificateAdd: undefined,
            openDocDialog: false,
            fileNames: [],
            downloadingTemplate: false,
            selectedAgreementToRequestCertificate: undefined,
            certificateRequestDialogIsOpen: false,
            sendingTheCertificateRequest: false,
            certificateRequestFailed: false,
            errorMessageInTheCertificateRequest: ""
        }
    }

    public async componentDidMount(): promise<void> {
        await this.onInit(true);
    }

    private async onInit(refresh?: boolean): promise<void> {
        const { spService, context, bkService, event } = this.props;
        const isArchived: boolean = event.StatusId === EventStatus.Archived;
        let { agreementStatuses, agreements, pendingCertificateAgreementsIds } = this.state;
        let [itemTypes, agreementsRes, pendingCertificateAgreements]: [ITermInfo[] | Term[], IEventAgreement[], string[]] = [[], [], []];
        if (refresh) {
            this.setState({ isLoading: true });
            try {
                [itemTypes, agreementsRes, pendingCertificateAgreements] = await promise.all([
                    spService.getTaxonomy(AgreementStatesTermSetId),
                    bkService.getAgreementsByEventId(event.BodyId, event.Id, isArchived),
                    bkService.getPendingCertificateAgreementsByEventId(event.BodyId, event.Id)
                ]);
                agreementStatuses = itemTypes.map((tag: ITermInfo) => Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName));
            }
            catch (error) {
                this.setState({ isLoading: false });
                Logger.error("Error MeetingsAgreements - OnInit - taxonomy and agreements initialization", error, context);
            }
        }
        else {
            agreementsRes = agreements;
            pendingCertificateAgreements = pendingCertificateAgreementsIds;
        }
        const docs: string[] = [];
        agreementsRes.forEach(item => {
            if (pendingCertificateAgreements.includes(item.Id)) {
                item.PendingCertificateTask = true;
            }
            item.RelatedDocumentsIds.forEach(docId => docs.push(docId));
            if (item.RelatedCertificateId && item.RelatedCertificateId !== "") {
                docs.push(item.RelatedCertificateId);
                item.RelatedDocumentsIds.push(item.RelatedCertificateId);
            }
        });
        let fileNames: IFileInfo[] = [];
        try {
            fileNames = await promise.all(docs.map(doc => spService.getFileInfoById(doc)));
        }
        catch (error) {
            Logger.error("Error MeetingsAgreements - OnInit - file info initialization", error, context);
        }

        this.setState({
            isLoading: false,
            agreementStatuses,
            checkedAgreement: undefined,
            selectedAgreementEditMode: undefined,
            selectedAgreementAttachmentAdd: undefined,
            selectedAgreementCertificateAdd: undefined,
            agreements: agreementsRes,
            pendingCertificateAgreementsIds: pendingCertificateAgreements,
            fileNames,
        });
    }

    private isFinalState(agreementState: string): boolean {
        return (agreementState === ApprodvedStatusId || agreementState === ApprodvedUnanimouslyStatusId || agreementState === DeclinedStatusId || agreementState === OverTheTableStatusId)
    }

    private setSelectedAgreementState(newAgreementState: string): void {
        const { selectedAgreementEditMode } = this.state;
        if (selectedAgreementEditMode) {
            this.setState({
                selectedAgreementEditMode: {
                    ...selectedAgreementEditMode,
                    StatusId: newAgreementState
                }
            });
        }
    }

    private setAgreementEditMode(selectedItem: IEventAgreement): void {
        this.setState({ selectedAgreementEditMode: selectedItem });
    }
    /*
    private setAgreementAttachmentMode(selectedItem: IEventAgreement): void {
        this.setState({ selectedAgreementAttachmentAdd: selectedItem, openDocDialog: true });
    }
    */
    private setAgreementCertificateMode(selectedItem: IEventAgreement): void {
        this.setState({ selectedAgreementCertificateAdd: selectedItem, openDocDialog: true });
    }

    // private generateDictionaryOrder(agreements: IEventAgreement[], items: IEventAgendaItem[]): { [key: string]: string } {

    //     const agendaItemsDictionary = items.reduce((memo: { [key: string]: IEventAgendaItem }, item: IEventAgendaItem) => {
    //         return { ...memo, [item.ItemSharedId]: item }
    //     }, {});

    //     const calculateOrder = (itemSharedId: string): string => {
    //         let order = '';
    //         let currentItem = agendaItemsDictionary[itemSharedId];
    //         let type = (agendaItemsDictionary[currentItem?.ParentId || '']?.OrderType ?? OrderTypes.Numeric) as OrderTypes

    //         while (currentItem) {
    //             order = `${getOrderFormatted(currentItem.Order - 1, type)}.${order}`;
    //             currentItem = agendaItemsDictionary[currentItem.ParentId || ''];
    //             type = (agendaItemsDictionary[currentItem?.ParentId || '']?.OrderType ?? OrderTypes.Numeric) as OrderTypes
    //         }

    //         // Remove the trailing dot and return the order
    //         return order.slice(0, -1);
    //     }



    //     const agreementsDictionary = agreements.reduce((memo: { [key: string]: string }, item: IEventAgreement) => {
    //         return { ...memo, [item.ItemSharedId]: calculateOrder(item.ItemSharedId) }

    //     }, {})
    //     return agreementsDictionary;
    // }

    private handleTitleBlur(value: string): void {
        const { selectedAgreementEditMode } = this.state;
        if (selectedAgreementEditMode) {
            this.setState({
                selectedAgreementEditMode: {
                    ...selectedAgreementEditMode,
                    Title: value
                }
            });
        }
    }

    private handleDescriptionBlur(value: string): void {
        const { selectedAgreementEditMode } = this.state;
        if (selectedAgreementEditMode) {
            this.setState({
                selectedAgreementEditMode: {
                    ...selectedAgreementEditMode,
                    Description: value
                }
            });
        }
    }

    private cancelAgreementChanges(): void {
        this.setState({
            selectedAgreementEditMode: undefined,
            selectedAgreementAttachmentAdd: undefined,
            selectedAgreementCertificateAdd: undefined
        });
    }

    private async saveAgreementChanges(): promise<void> {
        const { bkService, event } = this.props;
        const { agreements, selectedAgreementEditMode } = this.state;

        this.setState({ isLoadingItemId: selectedAgreementEditMode?.Id });

        if (selectedAgreementEditMode) {
            try {
                const updatedAgreement: IEventAgreement[] = await bkService.addOrUpdateEventAgreements(event.BodyId, event.Id, [selectedAgreementEditMode]);
                this.setState({
                    isLoadingItemId: undefined,
                    agreements: agreements?.map((agreement: IEventAgreement): IEventAgreement =>
                        (agreement.Id === updatedAgreement[0].Id) ? { ...updatedAgreement[0] } : agreement
                    ),
                    checkedAgreement: undefined,
                    selectedAgreementEditMode: undefined,
                    selectedAgreementAttachmentAdd: undefined,
                    selectedAgreementCertificateAdd: undefined
                });

                await this.onInit(true);
            }
            catch (error) {
                Logger.error("Error MeetingsAgreements - saveAgreementChanges - saving agreements", error, this.context);
                this.setState({ isLoadingItemId: undefined });
            }
        }
    }

    private async downloadCertification(): promise<void> {
        const { event, bkService } = this.props;
        const { checkedAgreement } = this.state;
        this.setState({ downloadingTemplate: true });
        try {
            if (checkedAgreement) {
                let result: ArrayBuffer = await bkService.downloadAgreementsTemplate(event.BodyId, event.Id, checkedAgreement.Id.toString());
                if (result) {
                    const url = window.URL.createObjectURL(new Blob([result]));
                    const enlace = document.createElement('a');
                    enlace.href = url;
                    enlace.setAttribute('download', `${strings.Certification} - ${event.Title.substring(0, 50)} - ${checkedAgreement.Title.substring(0, 25)}.docx`);
                    document.body.appendChild(enlace);
                    enlace.click();
                    document.body.removeChild(enlace);
                    window.URL.revokeObjectURL(url);
                }
            }
            this.setState({ downloadingTemplate: false });
        }
        catch (error) {
            Logger.error("Error MeetingsAgreements - downloadCertification - downloading template", error, this.context);
        }
    }

    private async removeRelatedDoc(agreementToEdit: IEventAgreement, uniqueId: string): promise<void> {
        const { event } = this.props;
        let { agreements } = this.state;
        // Start loading component
        this.setState({ isLoadingItemId: agreementToEdit.Id });
        try {
            await this.props.bkService.removeRelatedDocumentIdToAgreement(event.BodyId, event.Id, agreementToEdit.Id.toString(), uniqueId);
            agreements = agreements?.map(item => {
                if (item.Id === agreementToEdit.Id) {
                    return {
                        ...item,
                        RelatedDocumentsIds: item.RelatedDocumentsIds.filter(relatedDoc => relatedDoc !== uniqueId),
                        RelatedCertificateId: item.RelatedDocumentsIds.find(relatedCert => relatedCert === uniqueId) ? '' : item.RelatedCertificateId,
                    };
                }
                return item;
            });
            // Start loading component
            this.setState({ isLoadingItemId: undefined, agreements });
        }
        catch (error) {
            Logger.error("Error MeetingsAgreements - removeRelatedDoc - saving agreements", error, this.context);
            void this.onInit(true);
        }
    }

    private renderRelatedDocumentation(agendaItem: IEventAgreement): JSX.Element[] {
        const { context, isEditor, readOnly } = this.props;
        const { isLoading, selectedAgreementEditMode, fileNames } = this.state;
        const locked = !isEditor || readOnly;
        return agendaItem.RelatedDocumentsIds.map(file => {
            const document = find(fileNames, fileName => fileName.UniqueId === file);
            const blockDismiss: boolean = isLoading || selectedAgreementEditMode !== undefined;
            return (
                <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                    {document &&
                        <>
                            <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                <File className={styles.file} fileDetails={{ name: document?.Name, webUrl: document?.ServerRelativeUrl }} onClick={() => window.open(document?.ServerRelativeUrl + "?web=1", '_blank')} view={ViewType.oneline}></File>
                            </div>
                            <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                <ArrowDownload16Regular
                                    style={{ cursor: 'pointer' }}
                                    onClick={() => window.open(`${context.pageContext.site.serverRelativeUrl}/_layouts/download.aspx?SourceUrl=${document?.ServerRelativeUrl}`, '_self')}
                                />
                            </div>
                            {!locked &&
                                <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                    <ConfirmAction
                                        dialogTitle={strings.DeleteAssociation}
                                        dialogContent={<div>{strings.DeleteAssociatedDoc}<strong>{document?.Name}</strong> {strings.FromTheOrderOfTheDay} <strong>{agendaItem.Title}</strong>?</div>}
                                        dialogTrigger={<Dismiss16Regular style={{ flexShrink: 0, cursor: 'pointer', display: blockDismiss ? 'none' : 'block' }} />}
                                        acceptButtonprops={{
                                            appearance: 'primary',
                                            title: strings.Delete,
                                            icon: < DeleteRegular />,
                                            onClick: () => this.removeRelatedDoc(agendaItem, file)
                                        }}
                                        cancelButtonprops={{ appearance: 'secondary', title: strings.Cancel }} />
                                </div>
                            }
                        </>
                    }
                </div>
            );
        });
    }

    public render(): React.ReactElement<IMeetingAgreementsprops> {
        const {
            isEditor,
            isGuest,
            readOnly,
            context,
            spService,
            event,
            isOCprodle
        } = this.props;
        const {
            isLoading,
            isLoadingItemId,
            agreements,
            agreementStatuses,
            checkedAgreement,
            selectedAgreementEditMode,
            selectedAgreementAttachmentAdd,
            selectedAgreementCertificateAdd,
            openDocDialog,
            downloadingTemplate,
            selectedAgreementToRequestCertificate,
            certificateRequestDialogIsOpen,
            sendingTheCertificateRequest,
            certificateRequestFailed,
            errorMessageInTheCertificateRequest
        } = this.state;

        const locked: boolean = (!isEditor || readOnly);

        const ApprodvedStatusLabel: string | undefined = agreementStatuses && getTermLabel(agreementStatuses, ApprodvedStatusId);
        const ApprodvedUnanimouslyStatusLabel: string | undefined = agreementStatuses && getTermLabel(agreementStatuses, ApprodvedUnanimouslyStatusId);
        const InprodgressStatusLabel: string | undefined = agreementStatuses && getTermLabel(agreementStatuses, InprodgressStatusId);
        const excludedStatusLabels: string[] = [InprodgressStatusLabel, strings.Pending];
        return (
            <section className={styles.agreementsContainer}>
                {isLoading && isLoadingItemId === undefined ?
                    <Spinner style={{ flex: 1 }} label={strings.LoadingTheAgreements + "..."} />
                    :
                    (agreements && (agreements.length > 0)) ?
                        <>
                            <div className={styles.containerWrapper}>
                                {
                                    agreements.map((agreement: IEventAgreement, index: number): JSX.Element | undefined => {
                                        const agreementStatusLabel: string | undefined = ((agreement.StatusId === PendingStatusId) ?
                                            strings.Pending
                                            :
                                            (agreementStatuses && getTermLabel(agreementStatuses, agreement.StatusId))
                                        );
                                        if (agreementStatusLabel) {
                                            return (
                                                <div className={styles.itemRow} key={index} ref={(node: HTMLDivElement | null): void => {
                                                    // eslint-disable-next-line no-unused-extestssions
                                                    (selectedAgreementEditMode && selectedAgreementEditMode?.Id === agreement.Id) ?
                                                        node?.scrollIntoView({
                                                            behavior: "smooth",
                                                            block: "nearest"
                                                        })
                                                        :
                                                        null
                                                }}
                                                >
                                                    {
                                                        !locked &&
                                                        <div className={styles.itemColumn}>
                                                            <Radio
                                                                checked={checkedAgreement?.Id === agreement.Id}
                                                                onChange={(): void => this.setState({ checkedAgreement: agreement })}
                                                                style={{ visibility: this.isFinalState(agreement.StatusId) ? 'visible' : 'hidden' }}
                                                            />
                                                        </div>
                                                    }
                                                    <div className={`${styles.itemColumnGrow} ${styles.columnBlueBackground} ${styles.sidesPaddingAdjust}`}>
                                                        <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`} style={{ minHeight: "32px" }}>
                                                            <div className={`${styles.itemColumn} ${styles.justifyCenter}`}>
                                                                <Badge className={styles.orderCounter}>
                                                                    {agreement.Order || 1}
                                                                </Badge>
                                                            </div>
                                                            {
                                                                /* Estado de un acuerdo */
                                                                (selectedAgreementEditMode && selectedAgreementEditMode?.Id === agreement.Id) ?
                                                                    agreementStatuses?.map((agreementStatus: Term): JSX.Element | undefined => {
                                                                        if (!excludedStatusLabels.includes(agreementStatus.text)) {
                                                                            return (
                                                                                <Checkbox
                                                                                    label={agreementStatus.text}
                                                                                    value={agreementStatus.key}
                                                                                    disabled={isLoadingItemId !== undefined}
                                                                                    checked={selectedAgreementEditMode.StatusId === agreementStatus.key}
                                                                                    onChange={(): void => this.setSelectedAgreementState(agreementStatus.key)}
                                                                                />
                                                                            );
                                                                        }
                                                                    })
                                                                    :
                                                                    <div className={`${styles.itemColumnGrow} ${styles.sidesPaddingAdjust} ${styles.justifyCenter}`}>
                                                                        <div className={styles.itemRow}>
                                                                            {
                                                                                (agreementStatusLabel && !excludedStatusLabels.includes(agreementStatusLabel)) &&
                                                                                <Badge appearance="outline">
                                                                                    {
                                                                                        (agreementStatusLabel === ApprodvedUnanimouslyStatusLabel) ?
                                                                                            ApprodvedStatusLabel
                                                                                            :
                                                                                            agreementStatusLabel
                                                                                    }
                                                                                </Badge>
                                                                            }
                                                                        </div>
                                                                    </div>
                                                            }
                                                            {
                                                                !locked &&
                                                                (selectedAgreementEditMode?.Id !== agreement.Id) &&
                                                                <>
                                                                    {isLoadingItemId === agreement.Id && <Spinner size='tiny' />}
                                                                    <Button
                                                                        appearance="secondary"
                                                                        title={strings.Edit}
                                                                        icon={<EditRegular />}
                                                                        disabled={selectedAgreementEditMode !== undefined}
                                                                        onClick={(): void => this.setAgreementEditMode({ ...agreement })}
                                                                    />
                                                                    {/*
                                                                    <Button 
                                                                        appearance="secondary"
                                                                        title={strings.AssociateDocument}
                                                                        icon={<AttachRegular />}
                                                                        disabled={selectedAgreementEditMode !== undefined}
                                                                        onClick={(): void => this.setAgreementAttachmentMode({ ...agreement })}
                                                                    />
                                                                    */}
                                                                    {isECNTy(agreement.RelatedCertificateId) && this.isFinalState(agreement.StatusId) &&
                                                                        <Button
                                                                            appearance="secondary"
                                                                            title={strings.AssociateSignedCertificate}
                                                                            icon={<DocumentRibbonRegular />}
                                                                            disabled={selectedAgreementEditMode !== undefined}
                                                                            onClick={(): void => this.setAgreementCertificateMode({ ...agreement })}
                                                                        />
                                                                    }
                                                                </>
                                                            }
                                                            {
                                                                (
                                                                    (!isEditor && !readOnly) &&
                                                                    !agreement.RelatedCertificateId &&
                                                                    !isGuest && !isOCprodle && this.isFinalState(agreement.StatusId) && event?.ShowMeetingToolUrl
                                                                ) &&
                                                                <ConfirmAction
                                                                    dialogprops={{ open: (selectedAgreementToRequestCertificate === agreement) && certificateRequestDialogIsOpen }}
                                                                    dialogTitle={strings.RequestCertificate}
                                                                    dialogContent={
                                                                        <Field style={{ display: "flex", rowGap: "10px", flexDirection: "column" }}>
                                                                            <span>
                                                                                {strings.CertificateRequestQuestion} <strong>{`${agreement.Order}. ${agreement.Title}`}</strong>?
                                                                            </span>
                                                                            {
                                                                                certificateRequestFailed &&
                                                                                <Field style={{ color: "red" }}>
                                                                                    <span> {errorMessageInTheCertificateRequest} </span>
                                                                                </Field>
                                                                            }
                                                                        </Field>
                                                                    }
                                                                    dialogTrigger={
                                                                        (agreement?.PendingCertificateTask) ?
                                                                            <Tooltip
                                                                                content={strings.CertificateAlreadyRequested}
                                                                                positioning='before'
                                                                                relationship='label'
                                                                                withArrow
                                                                            >
                                                                                <Button
                                                                                    appearance="secondary"
                                                                                    // title={strings.RequestCertificate}
                                                                                    icon={<DocumentRibbonRegular />}
                                                                                    disabled={true}
                                                                                    onClick={(): void =>
                                                                                        this.setState({
                                                                                            selectedAgreementToRequestCertificate: agreement,
                                                                                            certificateRequestDialogIsOpen: true,
                                                                                            sendingTheCertificateRequest: false,
                                                                                            certificateRequestFailed: false,
                                                                                            errorMessageInTheCertificateRequest: ""
                                                                                        })
                                                                                    }
                                                                                />
                                                                            </Tooltip>
                                                                            :
                                                                            <Button
                                                                                appearance="secondary"
                                                                                title={strings.RequestCertificate}
                                                                                icon={<DocumentRibbonRegular />}
                                                                                disabled={selectedAgreementEditMode !== undefined}
                                                                                onClick={(): void =>
                                                                                    this.setState({
                                                                                        selectedAgreementToRequestCertificate: agreement,
                                                                                        certificateRequestDialogIsOpen: true,
                                                                                        sendingTheCertificateRequest: false,
                                                                                        certificateRequestFailed: false,
                                                                                        errorMessageInTheCertificateRequest: ""
                                                                                    })
                                                                                }
                                                                            />
                                                                    }
                                                                    loadingActionContent={
                                                                        sendingTheCertificateRequest && <Spinner size='tiny' label={strings.Requesting + "..."} />
                                                                    }
                                                                    acceptButtonprops={{
                                                                        appearance: 'primary',
                                                                        title: strings.Request,
                                                                        disabled: sendingTheCertificateRequest || certificateRequestFailed,
                                                                        onClick: (): promise<void> => this.requestCertificate(agreement.Id, agreement.Title)
                                                                    }}
                                                                    cancelButtonprops={{
                                                                        appearance: 'secondary',
                                                                        title: strings.RuleOut,
                                                                        disabled: sendingTheCertificateRequest,
                                                                        onClick: (): void => this.setState({
                                                                            selectedAgreementToRequestCertificate: undefined,
                                                                            certificateRequestDialogIsOpen: false
                                                                        })
                                                                    }}
                                                                />
                                                            }
                                                        </div>
                                                        <div className={`${styles.itemRow} ${styles.itemTitle} ${styles.topPaddingAdjust}`}>
                                                            <div className={`${styles.itemColumnGrow}`}>
                                                                {
                                                                    (selectedAgreementEditMode?.Id === agreement.Id) ?
                                                                        <Input
                                                                            type="text"
                                                                            name="Title"
                                                                            maxLength={255}
                                                                            placeholder={strings.WriteATitle}
                                                                            defaultValue={agreement.Title}
                                                                            disabled={isLoadingItemId !== undefined}
                                                                            onBlur={(e: React.FocusEvent<HTMLInputElement, Element>): void => this.handleTitleBlur(e.target.value)}
                                                                        />
                                                                        :
                                                                        agreement.Title
                                                                }
                                                            </div>
                                                        </div>
                                                        <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                                                            <div className={`${styles.itemColumnGrow} ${styles.itemDesc}`}>
                                                                {
                                                                    (selectedAgreementEditMode?.Id === agreement.Id) ?
                                                                        <Textarea
                                                                            className={styles.inputDesc}
                                                                            placeholder={strings.WriteADescription}
                                                                            defaultValue={agreement.Description ?? ""}
                                                                            disabled={isLoadingItemId !== undefined}
                                                                            onBlur={(e: React.FocusEvent<HTMLTextAreaElement, Element>): void => this.handleDescriptionBlur(e.target.value)}
                                                                        />
                                                                        :
                                                                        agreement.Description
                                                                }
                                                            </div>
                                                        </div>
                                                        {
                                                            selectedAgreementEditMode?.Id !== agreement.Id &&
                                                            <div className={`${styles.itemRow} ${styles.topPaddingAdjust}`}>
                                                                <div className={`${styles.itemColumnGrow}`}>
                                                                    {this.renderRelatedDocumentation(agreement)}
                                                                </div>
                                                            </div>
                                                        }
                                                        <div className={`${styles.itemRow} ${styles.topPaddingAdjust} ${styles.justifyEnd}`}>
                                                            {
                                                                (selectedAgreementEditMode && selectedAgreementEditMode?.Id === agreement.Id) &&
                                                                <>
                                                                    {isLoadingItemId && <Spinner size="extra-tiny" label={strings.Saving + "..."} />}
                                                                    <Button
                                                                        appearance="secondary"
                                                                        title={strings.Cancel}
                                                                        disabled={isLoadingItemId !== undefined}
                                                                        onClick={(): void => this.cancelAgreementChanges()}
                                                                    >
                                                                        {strings.Cancel}
                                                                    </Button>
                                                                    <Button
                                                                        appearance="primary"
                                                                        title={strings.Save}
                                                                        disabled={isLoadingItemId !== undefined}
                                                                        onClick={(): promise<void> => this.saveAgreementChanges()}
                                                                    >
                                                                        {strings.Save}
                                                                    </Button>
                                                                </>
                                                            }
                                                        </div>
                                                    </div>
                                                </div>
                                            );
                                        }
                                    })
                                }
                            </div>
                            {
                                !locked &&
                                <div className={`${styles.itemRow} ${styles.bottomButtonRow}`}>
                                    {downloadingTemplate && <Spinner size='tiny' />}
                                    <div className={`${styles.itemColumn}`}>
                                        <Button
                                            appearance="secondary"
                                            title={strings.DownloadCertification}
                                            icon={<ArrowDownload24Regular />}
                                            disabled={isLoadingItemId != undefined || this.state.checkedAgreement == undefined || downloadingTemplate}
                                            onClick={(): promise<void> => this.downloadCertification()}
                                        >
                                            {strings.DownloadCertification}
                                        </Button>
                                    </div>
                                </div>
                            }
                        </>
                        :
                        <span className={styles.noAgreementsAvailable}> {strings.NoAgreementsAvailable} </span>
                }
                {openDocDialog &&
                    (selectedAgreementAttachmentAdd ?
                        <NewDocForm
                            open={openDocDialog}
                            openCloseForm={this.onOpenCloseDocForm}
                            context={context}
                            spService={spService}
                            relativeUrlToGetDocs={event.StorageServerRelativeUrl}
                            relatedDocumentsIds={selectedAgreementAttachmentAdd?.RelatedDocumentsIds}
                            relatedItemTitle={`(${selectedAgreementAttachmentAdd?.Order}) ${selectedAgreementAttachmentAdd?.Title}`}
                            onSaveData={this.onSaveRelatedDocument}
                        />
                        :
                        <NewDocForm
                            open={openDocDialog}
                            openCloseForm={this.onOpenCloseDocForm}
                            context={context}
                            spService={spService}
                            relativeUrlToGetDocs={event.StorageServerRelativeUrl}
                            relatedDocumentsIds={selectedAgreementCertificateAdd ? selectedAgreementCertificateAdd.RelatedDocumentsIds : []}
                            relatedItemTitle={`(${selectedAgreementCertificateAdd?.Order}) ${selectedAgreementCertificateAdd?.Title}`}
                            onSaveData={this.onSaveRelatedCertificate}
                        />)
                }
            </section>
        );
    }

    private onSaveRelatedDocument = async (fileUniqueId: string): promise<void> => {
        const { bkService, event } = this.props;
        const { selectedAgreementAttachmentAdd } = this.state;
        let agreements = this.state.agreements;

        try {
            await bkService.addRelatedDocumentIdToAgreement(event.BodyId, event.Id, selectedAgreementAttachmentAdd ? selectedAgreementAttachmentAdd.Id.toString() : '', fileUniqueId, false);

            agreements = agreements.map(item => {
                if (item.Id === selectedAgreementAttachmentAdd?.Id) {
                    return {
                        ...item,
                        RelatedDocumentsIds: [...item.RelatedDocumentsIds, fileUniqueId],
                    };
                }

                return item;
            });

            this.setState({ isLoadingItemId: undefined, agreements });
            void this.onInit();
        } catch (error) {
            Logger.error("Error MeetingsAgreements - onSaveRelatedDocument - saving agreements", error, this.context);
            void this.onInit(true);
        }
    };

    private onSaveRelatedCertificate = async (fileUniqueId: string): promise<void> => {
        const { bkService, event } = this.props;
        const { selectedAgreementCertificateAdd } = this.state;
        let agreements = this.state.agreements;
        try {
            await bkService.addRelatedDocumentIdToAgreement(event.BodyId, event.Id, selectedAgreementCertificateAdd ? selectedAgreementCertificateAdd.Id.toString() : '', fileUniqueId, true);
            agreements = agreements.map(item => {
                if (item.Id === selectedAgreementCertificateAdd?.Id) {
                    return {
                        ...item,
                        RelatedCertificateId: fileUniqueId,
                    };
                }
                return item;
            });
            this.setState({ isLoadingItemId: undefined, agreements, isLoading: true });
            void this.onInit(true);
        }
        catch (error) {
            Logger.error("Error MeetingsAgreements - onSaveRelatedCertificate - saving agreements", error, this.context);
            this.setState({ isLoading: true });
            void this.onInit(true);
        }
    };

    private onOpenCloseDocForm = (): void => {
        this.setState({
            openDocDialog: !this.state.openDocDialog,
            selectedAgreementAttachmentAdd: undefined,
            selectedAgreementCertificateAdd: undefined
        });
    }

    private requestCertificate = async (agreementId: string, agreementTitle: string): promise<void> => {
        const { bkService, event } = this.props;
        const taskCertification: ITaskCertification = {
            AgreementSharedId: agreementId,
            SharedEventId: event.Id,
            Comment: "Comentario del certificado",
            AgreementTitle: agreementTitle
        }
        this.setState({ sendingTheCertificateRequest: true });
        try {
            await bkService.requestCertificate(event.BodyId, taskCertification);
            this.setState({
                selectedAgreementToRequestCertificate: undefined,
                certificateRequestDialogIsOpen: false
            });
            void this.onInit(true);
        }
        catch (error) {
            console.error(error);
            this.setState({
                sendingTheCertificateRequest: false,
                certificateRequestFailed: true,
                errorMessageInTheCertificateRequest: strings.BackendErrorMessage
            });
        }
    }
}