/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */

import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingStatusBar.module.scss';
import { IMeetingStatusBarprops, IMeetingStatusBarState as IMeetingStatusBarState } from './IMeetingStatusBar';
import {
    Edit16Filled,
    Location16Regular,
    VideoChat16Regular,
    SaveArrowRight20Regular,
    CalendarMail20Regular,
    Eye16Filled,
    Megaphone20Regular,
    Delete20Regular,
    CalendarCancel24Regular,
    CallEnd24Filled,
    CallForward20Regular,
    ArchiveArrowBack20Regular,
    CircleSmall20Regular
} from '@fluentui/react-icons';
import {
    Accordion,
    AccordionHeader,
    AccordionItem,
    AccordionPanel,
    Button,
    Field,
    Link,
    Menu,
    MenuButtonprops,
    MenuItem,
    MenuList,
    MenuPopover,
    MenuTrigger,
    Spinner,
    SplitButton,
    Textarea,
    Tooltip
} from '@fluentui/react-components';
import { Logger } from '../../../../utils/Logger';
import { TipoAlert } from '../../../../models/IAlert';
import { getTermLabel, getTimeFormatted } from '../../../../utils/Utils';
import { EventStatus } from '../../../../service/BackendServiceModels/EventModels';
import { INotification } from '../../../../service/BackendServiceModels/NotificationModel';
import MeetingForm from '../../../../components/MeetingForm';
import format from 'date-fns/format';
import es from 'date-fns/locale/es';
import ca from 'date-fns/locale/ca';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';
import ConfirmAction from '../../../../components/ConfirmAction';
import { IPendingChanges } from '../../../../models/IPendingChanges';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { ITermInfo } from '@pnp/sp/taxonomy';
import { Term } from '../../../../models/ITag';

const OnlineTerm: string = '8f4e0e22-2c3f-4454-8210-f4151dceeb89';
const OnlineInPersonTerm: string = 'e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f';
const AtendanceTypeSetId: string = '0775b014-e8fc-4aac-ae49-6b0b901156c1';
const OnlineToolSetId: string = '55dfa3e7-f0df-4dc8-b53c-4018418041fe';
const prodcEscritoTerm: string = '474ce74e-26c2-4ada-9dbb-516be5deb394';
const RemDocumentacionTerm: string = '4ab0be08-fef2-4537-ab14-3b3ae0025df3';
export default class MeetingStatusBar extends React.Component<IMeetingStatusBarprops, IMeetingStatusBarState> {

    private timer: any;
    private eventStatusToUpdate?: EventStatus;

    constructor(props: IMeetingStatusBarprops) {
        super(props);
        this.state = {
            openEditForm: false,
            updatesAvailable: false,

            sendUpdateDialogIsOpen: false,
            sendUpdateErrMsg: "",

            urgentNotifDialogOpen: false,
            urgentNotifErrMsg: "",
            textToSend: "",
            sendingUrgentNotif: false,

            dialogOpen: false,
            ErrMsg: "",
            nextState: undefined,

            pendingChanges: undefined,

            attendanceTypes: [],
            onlineTools: []
        }
    }

    async componentDidMount(): promise<void> {
        const { isEditor, readOnly, spService, context } = this.props;

        if (isEditor && !readOnly) {
            // Comprodbamos en la carga y activamos timer para recoger actualizaciones cada 30seg
            await this.checkUpdatesAvailable();
            this.timer = setInterval(() => this.checkUpdatesAvailable(), 30000);
        }

        try {
            const attendaceRes: ITermInfo[] = await spService.getTaxonomy(AtendanceTypeSetId);
            const attendanceTypes: Term[] = attendaceRes.map((tag: ITermInfo) => Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName));


            const onlineToolRes: ITermInfo[] = await spService.getTaxonomy(OnlineToolSetId);
            const onlineTools: Term[] = onlineToolRes.map((tag: ITermInfo) => Term.mapSPToTerm(tag, context.pageContext.cultureInfo.currentUICultureName));

            this.setState({ attendanceTypes, onlineTools });
        }
        catch (error) {
            console.log(error);
        }

    }

    componentWillUnmount(): void {
        if (this.timer) {
            clearInterval(this.timer);
        }
    }

    private async checkUpdatesAvailable(): promise<void> {
        const { bkService, event } = this.props;
        if (event) {
            const pendingChanges: IPendingChanges[] = await bkService.eventPendingChanges(event?.BodyId, event?.Id);
            this.setState({
                pendingChanges,
                updatesAvailable: pendingChanges && pendingChanges.length > 0
            });
        }
    }

    public openCloseFormFromParent(): void {
        this.setState({ openEditForm: !this.state.openEditForm });
    }

    public async changeStatus(nextStatusId: EventStatus): promise<void> {
        const { onStatusChange } = this.props;
        try {
            onStatusChange(nextStatusId);
            this.setState({ updatesAvailable: false });
        }
        catch (error) {
            Logger.error("Error MeetingStatus - openCloseFormFromParent - changing status", error, this.context);
        }
    }

    private handleTheButtonToRuleOutTheSendingUpdate = (): void => {
        this.setState({
            sendUpdateDialogIsOpen: false,
            sendUpdateErrMsg: ""
        });
    }

    private handleTheButtonToSendUpdate = async (): promise<void> => {
        const { event } = this.props;
        if (event) {
            try {
                if (this.eventStatusToUpdate) {
                    await this.changeStatus(this.eventStatusToUpdate);
                }
                this.setState({
                    sendUpdateDialogIsOpen: false
                });
            }
            catch (error) {
                console.error(error);
                this.setState({
                    sendUpdateErrMsg: strings.BackendErrorMessage
                });
            }
            finally {
                this.eventStatusToUpdate = undefined;
            }
        }
    }

    private handleTheButtonToRuleOutTheChangeStatus = (): void => {
        this.setState({ dialogOpen: false, ErrMsg: "" });
    }

    private handleTheButtonToChangeStatus = async (nextState: EventStatus): promise<void> => {
        const { event } = this.props;

        if (event) {
            try {
                await this.changeStatus(nextState);
                this.setState({
                    dialogOpen: false
                });
            }
            catch (error) {
                console.error(error);
                this.setState({
                    ErrMsg: strings.BackendErrorMessage
                });
            }
        }
    }


    private handleTheButtonToRuleOutTheSendingOfAnUrgentNotification = (): void => {
        this.setState({ urgentNotifDialogOpen: false, urgentNotifErrMsg: "", textToSend: "" });
    }

    private handleTheButtonToSendAnUrgentNotification = async (): promise<void> => {
        const { context, bkService, event } = this.props;
        const { textToSend } = this.state;

        if (textToSend !== "") {
            this.setState({ urgentNotifErrMsg: "", sendingUrgentNotif: true });
            try {
                // ToDo: Revisar las prodpiedades de INotification
                const urgentNotification: INotification = {
                    Id: "",
                    Title: event?.Title ?? "",
                    Body: textToSend,
                    Expires: event?.EndDate ?? new Date(),
                    Visualized: false,
                    Source: event?.BodyId ?? "",
                    Url: "",
                    Locale: context.pageContext.cultureInfo.currentUICultureName,
                    NotificationType: TipoAlert.Meetings,
                    NotificationPriority: ""
                };
                await bkService.addUrgentNotification(urgentNotification, event?.BodyId ?? "", event?.Id ?? "");
                this.setState({ urgentNotifErrMsg: "", textToSend: "", sendingUrgentNotif: false, urgentNotifDialogOpen: false });

            }
            catch (error) {
                console.error(error);
                this.setState({ urgentNotifErrMsg: strings.BackendErrorMessage });
            }
        }
        else {
            this.setState({ urgentNotifErrMsg: strings.MessageRequired });
        }
    }


    private filterPendingChanges(changeType: string): IPendingChanges[] | undefined {
        const { pendingChanges } = this.state;
        return pendingChanges?.filter((pendingChange: IPendingChanges): boolean => pendingChange.ChangeType === changeType);
    }

    public render(): React.ReactElement<IMeetingStatusBarprops> {
        const { context, event, loadingStatus, isEditor, loadingEventId, readOnly, isOCprodle } = this.props;
        const { updatesAvailable } = this.state;
        const locked = !isEditor || readOnly;
        const isEditable: boolean = !locked && event?.StatusId !== EventStatus.Canceled && event?.StatusId !== EventStatus.Archived;
        const locale = context.pageContext.cultureInfo.currentUICultureName;
        const locales: any = { 'es-ES': es, 'ca-ES': ca, 'eu-ES': eu, 'gl-ES': gl };
        const convoDate: string = event ? format(event.StartDate, 'PPPP', { locale: locales[locale] }) : "";
        const convoDateCap: string = convoDate && convoDate.charAt(0).toUpperCase() + convoDate.slice(1);
        return (
            <section className={styles.statusBar}>
                {event &&
                    <div className={styles.statusBarInfoContainer}>
                        <div className={styles.editZone}>
                            <div className={styles.infoRow}>
                                <Button
                                    size='small'
                                    appearance='primary'
                                    icon={isEditable ? <Edit16Filled /> : <Eye16Filled />}
                                    title={isEditable ? strings.EditTheMeetingDetails : strings.SeeTheMeetingDetails}
                                    aria-label={isEditable ? strings.EditTheMeetingDetails : strings.SeeTheMeetingDetails}
                                    onClick={(): void => this.openCloseFormFromParent()}
                                />
                            </div>
                            <div className={styles.infoRow}>
                                {convoDateCap && <span className={styles.infoRowItem}>{convoDateCap}</span>}
                                {event.StartDate && <span className={styles.infoRowItem}>{getTimeFormatted(event.StartDate)} - {getTimeFormatted(event.EndDate)}</span>}
                                {event.LocationDetails && <span className={styles.infoRowItem}>
                                    <Location16Regular className={styles.icon} />
                                    {event.LocationDetails}
                                </span>}
                                {
                                    (event.AttendanceTypeId === OnlineTerm || event.AttendanceTypeId === OnlineInPersonTerm) && event.ShowMeetingToolUrl && !isOCprodle &&
                                    (
                                        event.StatusId === EventStatus.InCelebration?
                                            <Link className={styles.infoRowItem} target='_blank' href={event.MeetingToolUrl}>
                                                <VideoChat16Regular /> {strings.Videoconference}
                                            </Link>
                                            :
                                            (event.StatusId === EventStatus.testBooking || event.StatusId === EventStatus.Published) &&
                                            <Tooltip relationship="label" content={strings.LinkAvailableSoon}>
                                                <Link className={styles.infoRowItem}>
                                                    <VideoChat16Regular /> {strings.Videoconference}
                                                </Link>
                                            </Tooltip>
                                    )
                                }
                            </div>
                        </div>
                        <div className={styles.titleAndActions}>
                            <div className={styles.meetingTitle}><span>{event.Title}</span></div>
                            {!locked &&
                                <div className={styles.actions}>
                                    {
                                        /* Botones para cambiar el estado de una Meeting */
                                        (loadingStatus && event.Id === loadingEventId) ?
                                            <Spinner size='extra-tiny' label={strings.Sending + "..."} />
                                            :
                                            {
                                                [EventStatus.InConstruction]:
                                                    <Menu positioning="below-end">
                                                        <MenuTrigger disableButtonEnhancement>
                                                            {
                                                                (triggerprops: MenuButtonprops): JSX.Element => (
                                                                    <SplitButton
                                                                        appearance="primary"
                                                                        icon={<SaveArrowRight20Regular />}
                                                                        menuButton={triggerprops}
                                                                        primaryActionButton={{
                                                                            onClick: (): void =>
                                                                                this.setState({
                                                                                    dialogOpen: true,
                                                                                    ErrMsg: "",
                                                                                    nextState: EventStatus.Published
                                                                                })
                                                                        }}
                                                                    >
                                                                        {strings.Publish}
                                                                    </SplitButton>
                                                                )
                                                            }
                                                        </MenuTrigger>
                                                        <MenuPopover>
                                                            <MenuList>
                                                                <MenuItem
                                                                    icon={<CalendarMail20Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.testBooking
                                                                        });
                                                                    }}

                                                                >
                                                                    {strings.ReserveAgenda}
                                                                </MenuItem>
                                                                <MenuItem
                                                                    icon={<Delete20Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.Canceled
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.DeleteMeeting}
                                                                </MenuItem>
                                                            </MenuList>
                                                        </MenuPopover>
                                                    </Menu>,
                                                [EventStatus.testBooking]:
                                                    <Menu positioning="below-end">
                                                        <MenuTrigger disableButtonEnhancement>
                                                            {
                                                                (triggerprops: MenuButtonprops): JSX.Element => (
                                                                    <SplitButton
                                                                        appearance='primary'
                                                                        icon={<SaveArrowRight20Regular />}
                                                                        menuButton={triggerprops}
                                                                        primaryActionButton={{
                                                                            onClick: (): void =>
                                                                                this.setState({
                                                                                    dialogOpen: true,
                                                                                    ErrMsg: "",
                                                                                    nextState: EventStatus.Published
                                                                                })

                                                                        }}
                                                                    >
                                                                        {strings.Publish}
                                                                    </SplitButton>
                                                                )
                                                            }
                                                        </MenuTrigger>
                                                        <MenuPopover>
                                                            <MenuList>
                                                                <Tooltip
                                                                    relationship="description"
                                                                    content={updatesAvailable ? strings.SendUpdatesToAllAttendees : strings.ThereAreNoUpdatesToSend}
                                                                >
                                                                    <MenuItem
                                                                        icon={<SaveArrowRight20Regular />}
                                                                        onClick={(): void => {
                                                                            this.eventStatusToUpdate = EventStatus.testBooking;
                                                                            this.setState({ sendUpdateDialogIsOpen: true });
                                                                        }}
                                                                        disabled={!updatesAvailable}
                                                                    >
                                                                        {strings.SendUpdate}
                                                                    </MenuItem>
                                                                </Tooltip>
                                                                <MenuItem
                                                                    icon={<CalendarCancel24Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.Canceled
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.CancelMeeting}
                                                                </MenuItem>
                                                            </MenuList>
                                                        </MenuPopover>
                                                    </Menu>,
                                                [EventStatus.Published]:
                                                    <Menu positioning="below-end">
                                                        <MenuTrigger disableButtonEnhancement>
                                                            {
                                                                (triggerprops: MenuButtonprops): JSX.Element => (
                                                                    <Tooltip
                                                                        relationship="description"
                                                                        content={updatesAvailable ? strings.SendUpdatesToAllAttendees : strings.ThereAreNoUpdatesToSend}
                                                                    >
                                                                        <SplitButton
                                                                            appearance='primary'
                                                                            icon={<SaveArrowRight20Regular />}
                                                                            menuButton={triggerprops}
                                                                            primaryActionButton={{
                                                                                onClick: (): void => {
                                                                                    this.eventStatusToUpdate = EventStatus.Published;
                                                                                    this.setState({ sendUpdateDialogIsOpen: true });
                                                                                },
                                                                                disabled: !updatesAvailable
                                                                            }}
                                                                        >
                                                                            {strings.SendUpdate}
                                                                        </SplitButton>
                                                                    </Tooltip>
                                                                )}
                                                        </MenuTrigger>
                                                        <MenuPopover>
                                                            <MenuList>
                                                                <MenuItem
                                                                    icon={<CallForward20Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.InCelebration
                                                                        });
                                                                    }}
                                                                >
                                                                    {event?.AttendanceTypeId === RemDocumentacionTerm? strings.StartMeetingRemDocumentacion : event?.AttendanceTypeId === prodcEscritoTerm ? strings.StartMeetingprodcEscrito :strings.StartMeeting}
                                                                </MenuItem>
                                                                <MenuItem
                                                                    icon={<CalendarCancel24Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.Canceled
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.CancelMeeting}
                                                                </MenuItem>
                                                                <MenuItem
                                                                    icon={<Megaphone20Regular />}
                                                                    onClick={(): void => { this.setState({ urgentNotifDialogOpen: true, urgentNotifErrMsg: "" }); }}
                                                                >
                                                                    {strings.UrgentNotification}
                                                                </MenuItem>
                                                            </MenuList>
                                                        </MenuPopover>
                                                    </Menu>,
                                                [EventStatus.InCelebration]:
                                                    <Menu positioning="below-end">
                                                        <MenuTrigger disableButtonEnhancement>
                                                            {
                                                                (triggerprops: MenuButtonprops): JSX.Element => (
                                                                    <Tooltip
                                                                        relationship="description"
                                                                        content={updatesAvailable ? strings.SendUpdatesToAllAttendees : strings.ThereAreNoUpdatesToSend}>
                                                                        <SplitButton
                                                                            appearance='primary'
                                                                            icon={<SaveArrowRight20Regular />}
                                                                            menuButton={triggerprops}
                                                                            primaryActionButton={{
                                                                                onClick: (): void => {
                                                                                    this.eventStatusToUpdate = EventStatus.InCelebration;
                                                                                    this.setState({ sendUpdateDialogIsOpen: true });
                                                                                },
                                                                                disabled: !updatesAvailable
                                                                            }}
                                                                        >
                                                                            {strings.SendUpdate}
                                                                        </SplitButton>
                                                                    </Tooltip>
                                                                )
                                                            }
                                                        </MenuTrigger>
                                                        <MenuPopover>
                                                            <MenuList>
                                                                <MenuItem
                                                                    icon={<Megaphone20Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            urgentNotifDialogOpen: true,
                                                                            urgentNotifErrMsg: ""
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.UrgentNotification}
                                                                </MenuItem>
                                                                <MenuItem
                                                                    icon={<CallEnd24Filled />}
                                                                    title={strings.FinishMeeting}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.Celebrated
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.FinishMeeting}
                                                                </MenuItem>
                                                            </MenuList>
                                                        </MenuPopover>
                                                    </Menu>,
                                                [EventStatus.Celebrated]:
                                                    <Menu positioning="below-end">
                                                        <MenuTrigger disableButtonEnhancement>
                                                            {
                                                                (triggerprops: MenuButtonprops): JSX.Element => (
                                                                    <Tooltip
                                                                        relationship="description"
                                                                        content={updatesAvailable ? strings.SendUpdatesToAllAttendees : strings.ThereAreNoUpdatesToSend}>
                                                                        <SplitButton
                                                                            appearance='primary'
                                                                            icon={<SaveArrowRight20Regular />}
                                                                            menuButton={triggerprops}
                                                                            primaryActionButton={{
                                                                                onClick: (): void => {
                                                                                    this.eventStatusToUpdate = EventStatus.Celebrated;
                                                                                    this.setState({ sendUpdateDialogIsOpen: true });
                                                                                },
                                                                                disabled: !updatesAvailable
                                                                            }}
                                                                        >
                                                                            {strings.SendUpdate}
                                                                        </SplitButton>
                                                                    </Tooltip>
                                                                )}
                                                        </MenuTrigger>
                                                        <MenuPopover>
                                                            <MenuList>
                                                                <MenuItem
                                                                    icon={<Megaphone20Regular />}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            urgentNotifDialogOpen: true,
                                                                            urgentNotifErrMsg: ""
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.UrgentNotification}
                                                                </MenuItem>
                                                                <MenuItem
                                                                    icon={<ArchiveArrowBack20Regular />}
                                                                    //onClick={async (): promise<void> => {await this.changeStatus(EventStatus.Archived)}}
                                                                    onClick={(): void => {
                                                                        this.setState({
                                                                            dialogOpen: true,
                                                                            ErrMsg: "",
                                                                            nextState: EventStatus.Archived
                                                                        });
                                                                    }}
                                                                >
                                                                    {strings.Archive}
                                                                </MenuItem>
                                                            </MenuList>
                                                        </MenuPopover>
                                                    </Menu>
                                            }[event.StatusId]
                                    }
                                </div>
                            }
                        </div>
                        {
                            this.state.openEditForm &&
                            <MeetingForm isEditor={isEditor} editMode={isEditor} readOnly={readOnly} eventData={this.props.event} dataChanged={this.props.onDataChanged.bind(this)}
                                bkService={this.props.bkService} spService={this.props.spService}
                                context={this.props.context} openForm={this.openCloseFormFromParent.bind(this)}
                            />
                        }
                    </div>
                }
                {this.renderTheDialogStatusChange()}
                {this.renderTheDialogToSendUpdate()}
                {this.renderTheDialogToSendAUrgentNotification()}

            </section >
        );
    }

    private renderTheDialogToSendUpdate(): JSX.Element {
        const { event, locale } = this.props;
        const { sendUpdateDialogIsOpen, attendanceTypes, onlineTools } = this.state
        const locales = { "es-ES": es, "ca-ES": ca, "eu-ES": eu, "gl-ES": gl };

        /* Filtrado de cambios en los datos básicos de la Meeting */
        const title: string | undefined = this.filterPendingChanges("Title")?.[0]?.ChangeValue;
        const eventTitle: string | undefined = title ?? event?.Title;
        const startDate: string | undefined = this.filterPendingChanges("StartDate")?.[0]?.ChangeValue;
        const endDate: string | undefined = this.filterPendingChanges("EndDate")?.[0]?.ChangeValue;
        const meetingType: string | undefined = this.filterPendingChanges("MeetingType")?.[0]?.ChangeValue;
        const location: string | undefined = this.filterPendingChanges("Location")?.[0]?.ChangeValue;
        const locationInformation: string | undefined = this.filterPendingChanges("LocationDetails")?.[0]?.ChangeValue;
        const onlineTool: string | undefined = this.filterPendingChanges("OnlineTool")?.[0]?.ChangeValue;
        const onlineToolUrl: string | undefined = this.filterPendingChanges("UrlOnlineTool")?.[0]?.ChangeValue;
        const description: string | undefined = this.filterPendingChanges("Description")?.[0]?.ChangeValue;

        /* Filtrado de cambios en el orden del día */
        const agendaPoints = this.filterPendingChanges("AgendaItem");

        /* Filtrado de cambios en los asistentes */
        const addedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("AddAttendance");
        const removedAssistants: IPendingChanges[] | undefined = this.filterPendingChanges("RemoveAttendance");

        /* Filtrado de cambios en los Documents */
        const documents: IPendingChanges[] | undefined = this.filterPendingChanges("Document");

        /* Filtrado de cambios en los acuerdos */
        const agreements: IPendingChanges[] | undefined = this.filterPendingChanges("Agreement");

        return (
            <>
                <ConfirmAction
                    dialogprops={{ open: sendUpdateDialogIsOpen }}
                    dialogTitle={strings.UpdateMeeting}
                    dialogContent={
                        <>
                            <Field style={{ display: "flex" }}>
                                <span>
                                    {strings.FirstPhraseToUpdateTheMeeting} <strong>{eventTitle}</strong> {strings.SecondPhraseToUpdateTheMeeting}
                                </span>
                            </Field>
                            <Accordion collapsible style={{ display: "flex", flexDirection: "column" }}>
                                {
                                    /* Cambios en los datos básicos de la Meeting */
                                    (title || location || locationInformation || meetingType || description || startDate || endDate || onlineTool || onlineToolUrl) && (
                                        <AccordionItem value="1">
                                            <AccordionHeader>
                                                <Field>
                                                    <span>
                                                        {strings.BasicData + ":"}
                                                    </span>
                                                </Field>
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    {
                                                        title &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewTitle}
                                                                </strong>
                                                                {": " + title}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        startDate &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewStartDate}
                                                                </strong>
                                                                {
                                                                    ": " +
                                                                    /* @ts-ignore */
                                                                    format(new Date(startDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                                }
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        endDate &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewEndDate}
                                                                </strong>
                                                                {
                                                                    ": " +
                                                                    /* @ts-ignore */
                                                                    format(new Date(endDate), 'PPPPp', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())
                                                                }
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        meetingType &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewMeetingType}
                                                                </strong>
                                                                {": " + getTermLabel(attendanceTypes, meetingType)}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        location &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewLocation}
                                                                </strong>
                                                                {": " + location}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        locationInformation &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewLocationInformation}
                                                                </strong>
                                                                {": " + locationInformation}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        onlineTool &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewOnlineTool}
                                                                </strong>
                                                                {": " + getTermLabel(onlineTools, onlineTool)}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        onlineToolUrl &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewOnlineToolUrl}
                                                                </strong>
                                                                {": " + onlineToolUrl}
                                                            </span>
                                                        </Field>
                                                    }
                                                    {
                                                        description &&
                                                        <Field style={{ display: "flex", alignItems: "flex-start" }}>
                                                            <CircleSmall20Regular style={{ minWidth: "20px" }} />
                                                            <span>
                                                                <strong>
                                                                    {strings.NewDescription}
                                                                </strong>
                                                                {": " + description.replace(/(<([^>]+)>)/ig, '')}
                                                            </span>
                                                        </Field>
                                                    }
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem >
                                    )
                                }
                                {
                                    /* Cambios en el orden del día */
                                    (agendaPoints && agendaPoints.length > 0) && (
                                        <AccordionItem value="2">
                                            <AccordionHeader>
                                                {strings.MeetingPoints + ":"}
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    <Field style={{ display: "flex", alignItems: "center" }}>
                                                        <CircleSmall20Regular />
                                                        {strings.TheOrderOfTheDayHasBeenUpdated}
                                                    </Field>
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem>
                                    )
                                }

                                {
                                    /* Adición de asistentes */
                                    (addedAssistants && addedAssistants.length > 0) && (
                                        <AccordionItem value="3">
                                            <AccordionHeader>
                                                {strings.AddedAssistants + ":"}
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    {
                                                        addedAssistants.map((addedAssistant: IPendingChanges): JSX.Element => (
                                                            <Field style={{ display: "flex", alignItems: "center" }}>
                                                                <CircleSmall20Regular />
                                                                <Person userId={addedAssistant.ChangeValue} view={ViewType.oneline} />
                                                            </Field>
                                                        ))
                                                    }
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem>
                                    )
                                }
                                {
                                    /* Eliminación de asistentes */
                                    (removedAssistants && removedAssistants.length > 0) && (
                                        <AccordionItem value="4">
                                            <AccordionHeader>
                                                {strings.RemovedAssistants + ":"}
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    {
                                                        removedAssistants.map((removedAssistant: IPendingChanges): JSX.Element => (
                                                            <Field style={{ display: "flex", alignItems: "center" }}>
                                                                <CircleSmall20Regular />
                                                                <Person userId={removedAssistant.ChangeValue} view={ViewType.oneline} />
                                                            </Field>
                                                        ))
                                                    }
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem>
                                    )
                                }
                                {
                                    /* Cambios en los Documents */
                                    (documents && documents.length > 0) && (
                                        <AccordionItem value="5">
                                            <AccordionHeader>
                                                {strings.Documents + ":"}
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    {
                                                        documents.map((document: IPendingChanges): JSX.Element => (
                                                            <Field style={{ display: "flex", alignItems: "center" }}>
                                                                <CircleSmall20Regular />
                                                                {document.ChangeValue}
                                                            </Field>
                                                        ))
                                                    }
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem>
                                    )
                                }
                                {
                                    /* Cambios en los acuerdos */
                                    (agreements && agreements.length > 0) && (
                                        <AccordionItem value="6">
                                            <AccordionHeader>
                                                {strings.Agreements + ":"}
                                            </AccordionHeader>
                                            <AccordionPanel style={{ display: "flex", flexDirection: "column" }}>
                                                <Field style={{ paddingLeft: "27px" }}>
                                                    {
                                                        agreements.map((agreement: IPendingChanges): JSX.Element => (
                                                            <Field style={{ display: "flex", alignItems: "center" }}>
                                                                <CircleSmall20Regular />
                                                                {agreement.ChangeValue}
                                                            </Field>
                                                        ))
                                                    }
                                                </Field>
                                            </AccordionPanel>
                                        </AccordionItem>
                                    )
                                }
                            </Accordion>
                        </>
                    }
                    acceptButtonprops={{
                        appearance: 'primary',
                        title: "Actualizar",
                        onClick: (): promise<void> => this.handleTheButtonToSendUpdate()
                    }}
                    cancelButtonprops={{
                        appearance: "secondary",
                        title: strings.RuleOut,
                        onClick: (): void => this.handleTheButtonToRuleOutTheSendingUpdate()
                    }}
                />
            </>
        );
    }

    private renderTheDialogStatusChange(): JSX.Element {
        const { event } = this.props;
        const {
            dialogOpen,
            ErrMsg,
            nextState
        } = this.state;

        let dialogText: string = "";
        let dialogContent: string = "";

        switch (nextState) {
            case EventStatus.Canceled:
                if (event?.StatusId === EventStatus.InConstruction) {
                    dialogText = strings.DeleteMeeting;
                    dialogContent = strings.QuestionToDeleteTheMeeting;
                } else {
                    dialogText = strings.CancelMeeting;
                    dialogContent = strings.QuestionToCancelTheMeeting;
                }
                break;
            case EventStatus.Published:
                dialogText = strings.PublishMeeting;
                dialogContent = strings.QuestionToPublishTheMeeting;
                break;
            case EventStatus.testBooking:
                dialogText = strings.testBookMeeting;
                dialogContent = strings.QuestionTotestBookTheMeeting;
                break;
            case EventStatus.InCelebration:
                dialogText = event?.AttendanceTypeId === prodcEscritoTerm ? strings.StartMeetingprodcEscrito : event?.AttendanceTypeId === RemDocumentacionTerm ? strings.StartMeetingRemDocumentacion : strings.StartMeeting;
                dialogContent = strings.QuestionToStartTheMeeting;
                break;
            case EventStatus.Celebrated:
                dialogText = strings.FinishMeeting;
                dialogContent = strings.QuestionToFinishTheMeeting;
                break;
            case EventStatus.Archived:
                dialogText = strings.ArchiveMeeting;
                dialogContent = strings.QuestionToArchiveTheMeeting;
                break;
        }


        return (
            <>
                {
                    (ErrMsg && ErrMsg.trim() !== "") &&
                    <span style={{ color: "red", marginTop: "10px" }}> {ErrMsg} </span>
                }
                <ConfirmAction
                    dialogprops={{ open: dialogOpen }}
                    dialogTitle={dialogText}
                    dialogContent={dialogContent}
                    acceptButtonprops={{
                        appearance: 'primary',
                        title: dialogText,
                        onClick: () => { if (nextState) { void this.handleTheButtonToChangeStatus(nextState) } }
                    }}
                    cancelButtonprops={{
                        appearance: "secondary",
                        title: strings.RuleOut,
                        onClick: () => this.handleTheButtonToRuleOutTheChangeStatus()
                    }} />
            </>
        );
    }



    private renderTheDialogToSendAUrgentNotification(): JSX.Element {
        const { urgentNotifDialogOpen, sendingUrgentNotif, urgentNotifErrMsg, textToSend } = this.state;

        return (
            <ConfirmAction
                dialogprops={{ open: urgentNotifDialogOpen }}
                dialogTitle={strings.UrgentNotification}
                dialogContent={
                    <>
                        <Field validationMessage={urgentNotifErrMsg} required>
                            <Textarea
                                required
                                style={{ height: "100px" }}
                                placeholder={strings.WriteAnUrgentNotification}
                                disabled={sendingUrgentNotif}
                                value={textToSend}
                                onChange={(ev, data) => {
                                    if (data.value.length > 160) {
                                        this.setState({ urgentNotifErrMsg: strings.MaximumCharactersForUrgentNotification });
                                    }
                                    else {
                                        if (data.value.length === 0) {
                                            this.setState({ textToSend: data.value, urgentNotifErrMsg: strings.MessageRequired });
                                        }
                                        else {
                                            this.setState({ textToSend: data.value, urgentNotifErrMsg: "" });
                                        }
                                    }
                                }}
                            />
                        </Field>
                    </>
                }
                loadingActionContent={sendingUrgentNotif && <Spinner size="extra-tiny" label={strings.Sending + "..."} />}
                acceptButtonprops={{
                    appearance: 'primary',
                    title: strings.Send,
                    disabled: sendingUrgentNotif,
                    onClick: () => this.handleTheButtonToSendAnUrgentNotification()
                }}
                cancelButtonprops={{
                    appearance: "secondary",
                    title: strings.RuleOut,
                    disabled: sendingUrgentNotif,
                    onClick: () => this.handleTheButtonToRuleOutTheSendingOfAnUrgentNotification()
                }} />
        );
    }



}