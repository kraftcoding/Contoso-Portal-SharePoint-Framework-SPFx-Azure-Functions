/* eslint-disable @typescript-eslint/ban-ts-comment */
import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './MeetingAssistants.module.scss';
import { IMeetingAssistantsprops, IMeetingAssistantsState } from './IMeetingAssistants';
import {
    Button,
    Checkbox,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTitle,
    Field,
    Spinner,
    Tooltip
} from '@fluentui/react-components';
import {
    AddRegular,
    ArrowDownloadRegular,
    ArrowEnterLeft20Regular,
    ArrowRight20Regular,
    // DesktopCursorRegular,
    // HomePersonRegular,
    DesktopCursor20Regular,
    Dismiss20Regular,
    HandLeft20Regular,
    HomePerson20Regular,
    Info20Regular,
    PeopleCheckmarkRegular,
    Person20Regular,
} from '@fluentui/react-icons';
import { EventStatus, IUserAttendance } from '../../../../service/BackendServiceModels/EventModels';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import { Logger } from '../../../../utils/Logger';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import { PersonType } from '@microsoft/mgt-spfx';
import { isECNTy } from '@microsoft/sp-lodash-subset';
import { sortByprodp } from '../../../../utils/Utils';
import ConfirmAction from '../../../../components/ConfirmAction';
import { getBodyIdFormUrl } from '../../../../service/RoleService';

const pendingTermId: string = 'b6d26a81-48bb-4a65-a0d6-7d50b3feed1a';
const acceptedTermId: string = '9628620a-c31b-4903-9bb2-b305d51821bd';
const refusedTermId: string = 'aac5817d-310f-4bfa-838a-1a1dda983bfb';
const automaticRefusedTermId: string = '8f52f3b1-90e6-429e-b796-6d6ad9f077b2';
const delegationPendingTermId: string = 'b90d7bd3-2859-4974-a0e1-f9808701bebd';
const delegatedTermId: string = 'fccf999f-fd9a-4370-a4a3-11400a5d19ff';
const onlineAndInPersonTermId: string = 'e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f';
const onlineTermId: string = '8f4e0e22-2c3f-4454-8210-f4151dceeb89';
const inPersonTermId: string = '3a16b292-5616-4ef9-8a56-2b36c20bb992';

export default class MeetingDocs extends React.Component<IMeetingAssistantsprops, IMeetingAssistantsState> {

    private testvAttendances: IUserAttendance[] = [];

    constructor(props: IMeetingAssistantsprops) {
        super(props);
        this.state = {
            attendances: [],
            isLoading: false,
            isSaving: false,
            openAdd: false,
            userError: '',
            editing: false,
            downloadingAttendance: false,
            groupId: '',
            reportingAttendance: false,
            reportingAttendanceErrorMessage: "",
            membersGroupId: "",
            guestsGroupId: "",
            convoGroupId:"",
            gestorGroupId: "",
            asistenteGroupId: "",
        };
    }

    public componentDidMount(): void {
        void this.onInit();
    }

    private async onInit(attendances?: IUserAttendance[]): promise<void> {
        const { bkService, event, context, spService } = this.props;
        const isArchived: boolean = event?.StatusId === EventStatus.Archived;

        this.setState({ isLoading: true });

        if (attendances) {
            attendances = attendances.filter(user => !isECNTy(user.UserPrincipalName));
            this.setState({ attendances: sortByprodp(attendances, 'UserPrincipalName'), isLoading: false, openAdd: false, isSaving: false, userError: '', loadingAttendance: undefined });
        }
        else {
            try {
                const bodyId: string = getBodyIdFormUrl(context.pageContext.site.serverRelativeUrl);
                let [groupId, membersGroupId,guestsGroupId,convoGroupId,gestorGroupId,asistenteGroupId, attendances]: [string,string, string, string, string, string, IUserAttendance[]] = await promise.all([
                    spService.getGroupId(bodyId), 
                    spService.getGroupId(`${bodyId}_members`),
                    spService.getGroupId(`${bodyId}_guests`),
                    spService.getGroupId(`${bodyId}_schedulers`),
                    spService.getGroupId(`${bodyId}_gestorschedulers`),
                    spService.getGroupId(`${bodyId}_asistentemembers`), 
                    bkService.getAttendanceByEventId(event.BodyId, event.Id, isArchived)])
                attendances = attendances.filter(user => !isECNTy(user.UserPrincipalName));
                this.setState({ attendances: sortByprodp(attendances, 'UserPrincipalName'), isLoading: false, openAdd: false, isSaving: false, userError: '', loadingAttendance: undefined, groupId, membersGroupId, convoGroupId, guestsGroupId, gestorGroupId, asistenteGroupId });
            }
            catch (error) {
                Logger.error("Error MeetingsAssistants - onInit", error, this.context);
                this.setState({ isLoading: false });
            }
        }
    }

    private renderDeleteBtn(user: IUserAttendance): JSX.Element {
        const { isLoading } = this.state;

        return (
            <ConfirmAction
                dialogTitle={strings.CancelAttendance}
                dialogContent={
                    <div className={styles.dialogCancelAttContent}>
                        <span>{strings.CancelAttendanceText}</span>
                        <Person userId={user.UserPrincipalName} view={ViewType.oneline} fetchImage={false} />
                    </div>
                }
                dialogTrigger={<div title={strings.CancelAttendance}>
                    <Dismiss20Regular className={styles.deleteButton} style={{ display: isLoading ? 'none' : 'block' }} />
                </div>}
                acceptButtonprops={{
                    appearance: 'primary',
                    title: strings.CancelAttendance,
                    onClick: () => this.CancelAttendance(user)
                }}
                cancelButtonprops={{ appearance: 'secondary', title: strings.Cancel }} />
        );
    }

    private async CancelAttendance(user: IUserAttendance): promise<void> {
        const { bkService, event } = this.props;
        let { attendances } = this.state;

        // Loading front before doing the back call
        const loadingAttendance = attendances.find(item => item.UserPrincipalName === user.UserPrincipalName)?.UserPrincipalName;

        this.setState({ loadingAttendance });

        try {
            attendances = await bkService.deleteAttendance(event.BodyId, event.Id, user.UserPrincipalName);

            void this.onInit(attendances);
        } catch (error) {
            Logger.error("Error MeetingsAssistants - CancelAttendance", error, this.context);
        }
    }

    private addAttendance = async (): promise<void> => {
        const { bkService, event } = this.props;
        const { attendances, form } = this.state;

        if (form?.UserPrincipalName && !isECNTy(form.UserPrincipalName)) {
            const currentUser: IUserAttendance | undefined = attendances.find(att => att.UserPrincipalName === form.UserPrincipalName);

            if (!currentUser) {
                this.setState({ isSaving: true, userError: '' });

                try {
                    const newAttendances: IUserAttendance[] = await bkService.addOrUpdateAttendance(event.BodyId, event.Id, form.UserPrincipalName);
                    this.setState({ openAdd: false, isSaving: false, userError: '' });

                    void this.onInit(newAttendances);
                } catch (error) {
                    Logger.error("Error MeetingsAssistants - addAttendance", error, this.context);
                    this.setState({ openAdd: false, isSaving: false, userError: '' });
                }
            } else {
                this.setState({ userError: strings.UserAlreadyAdded });
            }
        } else {
            this.setState({ userError: strings.UserRequired });
        }
    }

    private async updateAttendance(): promise<void> {
        const { bkService, event } = this.props;
        const { attendances } = this.state;

        this.setState({ reportingAttendance: true });

        try {
            await bkService.updateAttendance(event.BodyId, event.Id, attendances);
            this.setState({ reportingAttendance: false, editing: false });

        } catch (error) {
            Logger.error("Error MeetingsAssistants - updateAttendance", error, this.context);
            // void this.onInit();
            this.setState({ reportingAttendanceErrorMessage: strings.FailedToReportAssistance, reportingAttendance: false });
        }
    }

    private async downloadAttendance(): promise<void> {
        const { event, bkService } = this.props;
        this.setState({ downloadingAttendance: true });
        try {
            const result: ArrayBuffer = await bkService.downloadAttendanceTemplate(event.BodyId, event.Id);
            if (result) {
                const url = window.URL.createObjectURL(new Blob([result], { type: "application/pdf" }));
                const enlace = document.createElement('a');
                enlace.href = url;
                enlace.setAttribute('download', `${strings.Attendance} - ${event.Title.substring(0, 50)}.pdf`);
                document.body.appendChild(enlace);
                enlace.click();
                document.body.removeChild(enlace);
                window.URL.revokeObjectURL(url);
            }

            this.setState({ downloadingAttendance: false });
        } catch (error) {
            Logger.error("Error MeetingsAssistants - downloadAttendance", error, this.context);
            this.setState({ downloadingAttendance: false });
        }
    }

    private onDismiss = (): void => {
        this.setState({ openAdd: false, userError: '' });
    }

    private renderAddAssistantBtn(): JSX.Element {
        const { isSaving, openAdd, userError, form,gestorGroupId, guestsGroupId, membersGroupId, convoGroupId, asistenteGroupId } = this.state;
        
        return (
            <Dialog open={openAdd} onOpenChange={this.onDismiss}>
                <DialogSurface className={styles.dialogAssistantsSurface}>
                    <DialogBody>
                        <DialogTitle tabIndex={0}>{strings.AddAssistant}</DialogTitle>
                        <DialogContent className={styles.dialogAssistantsContent}>
                            <Field
                                validationMessage={userError}
                                style={{ paddingBottom: "5px" }}
                            >
                                <PeoplePicker
                                    type={PersonType.person}
                                    //userType={UserType.user}
                                    placeholder={strings.AddAUser}
                                    groupIds={[membersGroupId, convoGroupId, guestsGroupId, gestorGroupId, asistenteGroupId]}
                                    transitiveSearch={true}
                                    selectionMode='single'
                                    selectionChanged={async (e) => {
                                        e.detail.forEach((value: any, index, array) => this.setState({
                                            form: {
                                                isGuest: form?.isGuest,
                                                UserPrincipalName: value.userPrincipalName
                                            },
                                            userError: ''
                                        }));
                                    }}

                                    disabled={isSaving}

                                />
                            </Field>
                        </DialogContent>
                        <DialogActions>
                            {isSaving && <Spinner size='tiny' />}
                            <Button disabled={isSaving} appearance='secondary' onClick={this.onDismiss}>{strings.RuleOut}</Button>
                            <Button disabled={isSaving} appearance='primary' onClick={this.addAttendance}>{strings.Save}</Button>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

    private changeAttend(upn: string, checked: boolean): void {
        const { attendances } = this.state;

        this.setState({
            attendances: attendances.map(att => {
                if (att.UserPrincipalName === upn) {
                    return {
                        ...att,
                        HasAttended: checked
                    };
                } else {
                    return att;
                }
            })
        });
    }

    public render(): React.ReactElement<IMeetingAssistantsprops> {
        const { isEditor, event, readOnly } = this.props;
        const { attendances, isLoading, editing, downloadingAttendance, reportingAttendance, reportingAttendanceErrorMessage } = this.state;

        const locked = !isEditor || readOnly;

        return (
            <section className={styles.meetingAssistants}>
                {isLoading ?
                    <Spinner style={{ margin: 'auto' }} label={`${strings.LoadingAttendees}...`} />
                    :
                    (attendances.length === 0 ?
                        <div className={styles.noElementsContainer}>
                            <span>{strings.ThereAreNoAssistanceToShow}</span>
                        </div>
                        :
                        <div className={styles.meetingAssistantsPrimaryContainer}>
                            {this.renderGroup(pendingTermId, strings.Pending)}
                            {this.renderGroup(acceptedTermId, strings.Accepted)}
                            {this.renderGroup(refusedTermId, strings.Refused)}
                            {this.renderDelegations()}
                        </div>
                    )
                }
                {!locked && !isLoading &&
                    (!editing ?
                        <div className={styles.meetingAssistantsButtons}>
                            {
                                downloadingAttendance && <Spinner size='tiny' />
                            }
                            <Button
                                disabled={downloadingAttendance}
                                onClick={() => this.downloadAttendance()}
                                icon={<ArrowDownloadRegular />}
                            >
                                {strings.AttendeesList}
                            </Button>
                            {this.renderAddAssistantBtn()}
                            {/* Don't allow to change HasAttended until status is "En celebración" or "Celebrada" */}
                            {(event.StatusId === EventStatus.InCelebration || event.StatusId === EventStatus.Celebrated) &&
                                <Button
                                    icon={<PeopleCheckmarkRegular />}
                                    onClick={() => {
                                        this.testvAttendances = attendances;
                                        this.setState({ editing: true })
                                    }}
                                >
                                    {strings.ReportAttendance}
                                </Button>
                            }
                            <Button icon={<AddRegular />} appearance='primary' onClick={() => this.setState({ openAdd: true })}>{strings.AddAssistant}</Button>
                        </div>
                        :
                        <div className={styles.meetingAssistantsButtons}>
                            {
                                (reportingAttendance) ?
                                    <Spinner size='tiny' label={strings.Sending + "..."} />
                                    :
                                    (reportingAttendanceErrorMessage === "") ?
                                        <span className={styles.infoText}> <Info20Regular />{strings.CheckTheUsersWhoAttendedTheMeeting + "."}</span>
                                        :
                                        <span className={styles.infoText} style={{ color: "red" }}> <Info20Regular />{reportingAttendanceErrorMessage + "."}</span>
                            }
                            <div>
                                <Button disabled={
                                    reportingAttendance && reportingAttendanceErrorMessage === ""
                                }
                                    onClick={() => this.setState({
                                        editing: false,
                                        attendances: this.testvAttendances,
                                        reportingAttendanceErrorMessage: "",
                                        reportingAttendance: false
                                    })
                                    }>
                                    {strings.Cancel}
                                </Button>
                                <Button appearance='primary' disabled={
                                    reportingAttendance || reportingAttendanceErrorMessage !== ""
                                } onClick={() => this.updateAttendance()}>
                                    {strings.Save}
                                </Button>
                            </div>
                        </div>
                    )
                }
            </section>
        );
    }

    private handleOnlineAndInPerson = (AttendanceTypeId: string): { texto: string, icono: JSX.Element } => {
        switch (AttendanceTypeId) {
            case onlineTermId:
                return { texto: strings.Online, icono: <DesktopCursor20Regular className={styles.delegatedVote} /> };

            case inPersonTermId:
                return { texto: strings.InPerson, icono: <HomePerson20Regular className={styles.delegatedVote} /> };

            default:
                return { texto: "", icono: <></> };
        }
    }

    private renderGroup(termId: string, groupName: string): React.ReactElement | undefined {
        const { isEditor, readOnly, event } = this.props;
        const { attendances, editing, loadingAttendance } = this.state;
        let groupOfPeople: IUserAttendance[] = [];
        const locked = !isEditor || readOnly;
        if (termId === pendingTermId) {
            // Pendientes = Pendiente + Rechazo automático 
            groupOfPeople = attendances.filter(user => (user.RequestStatusId === pendingTermId || user.RequestStatusId === automaticRefusedTermId));
        }
        else {
            groupOfPeople = attendances.filter(user => user.RequestStatusId === termId);
        }
        if (groupOfPeople.length > 0)
            return (
                <div className={styles.meetingAssistantsSecondaryContainer}>
                    <span className={styles.meetingAssistantsTitle}>{groupName}</span>
                    <div className={styles.meetingAssistantsList}>
                        <div className={styles.meetingAssistantsPersonsContainer}>
                            {groupOfPeople.map((person: IUserAttendance, index): JSX.Element => {
                                const personAttendanceType: { texto: string, icono: JSX.Element } = this.handleOnlineAndInPerson(person.AttendanceTypeId);
                                return (
                                    <div key={index}>
                                        <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                            <Checkbox
                                                checked={person.HasAttended}
                                                disabled={!editing}
                                                onChange={(ev, data) => this.changeAttend(person.UserPrincipalName, data.checked as boolean)}
                                                style={!isEditor ? { visibility: 'hidden' } : {}}
                                            />
                                            <Person
                                                userId={person.UserPrincipalName}
                                                view={ViewType.threelines}
                                                line3property='officeLocation'
                                                className={styles.persona}
                                                fetchImage={false}
                                            />
                                            {
                                                /* Tipo de Attendance */
                                                (
                                                    (event.AttendanceTypeId === onlineAndInPersonTermId) &&
                                                    (groupName === strings.Accepted) &&
                                                    (personAttendanceType.texto !== "")
                                                ) &&
                                                <Tooltip content={personAttendanceType.texto} relationship='label' withArrow>
                                                    {personAttendanceType.icono}
                                                </Tooltip>
                                            }
                                            {!locked && !editing && (
                                                loadingAttendance === person.UserPrincipalName ?
                                                    <Spinner className={styles.deleteButton} size='tiny' />
                                                    :
                                                    this.renderDeleteBtn(person))
                                            }
                                        </div>
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                </div>
            );
    }

    private renderDelegations(): React.ReactElement | undefined {
        const { attendances, editing, loadingAttendance } = this.state;
        const { isEditor, readOnly } = this.props;
        const delegationPeople: IUserAttendance[] = attendances.filter(user => user.RequestStatusId === delegationPendingTermId || user.RequestStatusId === delegatedTermId);
        const locked = !isEditor || readOnly;

        if (delegationPeople.length > 0)
            return (
                <div className={styles.meetingAssistantsDelegatedContainer}>
                    <div className={styles.meetingAssistantsTitle}>
                        <span> {strings.Delegated} </span>
                    </div>
                    <div className={styles.meetingAssistantsPersonsContainer}>
                        {delegationPeople.map((delegationPerson: IUserAttendance): JSX.Element => {
                            return (
                                <div className={styles.meetingAssistants_alternativePrimaryContainer}>
                                    <div className={styles.meetingAssistants_alternativeSecondaryContainer}>
                                        <div>
                                            <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                                <Checkbox
                                                    checked={delegationPerson.HasAttended}
                                                    disabled={!this.state.editing}
                                                    onChange={(ev, data) => this.changeAttend(delegationPerson.UserPrincipalName, data.checked as boolean)}
                                                    style={{ visibility: 'hidden' }}
                                                />
                                                <Person
                                                    userId={delegationPerson.UserPrincipalName}
                                                    view={ViewType.threelines}
                                                    line3property='officeLocation'
                                                    className={styles.persona}
                                                    fetchImage={false}
                                                />
                                                {!locked && !editing && (
                                                    loadingAttendance === delegationPerson?.UserPrincipalName ?
                                                        <Spinner className={styles.deleteButton} size='tiny' />
                                                        :
                                                        this.renderDeleteBtn(delegationPerson))
                                                }
                                                <ArrowRight20Regular className={styles.meetingAssistants_rightArrow} />
                                            </div>
                                        </div>
                                    </div>
                                    {!isECNTy(delegationPerson.DelegationUserPrincipalName) ?
                                        <div className={styles.meetingAssistants_alternativeSecondaryContainer}>
                                            <div>
                                                <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                                    <ArrowEnterLeft20Regular className={styles.meetingAssistants_diagonalArrow} style={{ transform: 'scaleX(-1)' }} />
                                                    <Person
                                                        userId={delegationPerson.DelegationUserPrincipalName}
                                                        view={ViewType.threelines}
                                                        line3property='officeLocation'
                                                        className={styles.persona}
                                                        fetchImage={false}
                                                    />
                                                    {
                                                        /* Se delega la Attendance pero no el voto */
                                                        (
                                                            !delegationPerson.VoteDelegation &&
                                                            delegationPerson.DelegationUserPrincipalName !== ""
                                                        ) &&
                                                        <Tooltip content={strings.AttendanceDelegation} relationship='label' withArrow>
                                                            <Person20Regular className={styles.delegatedVote} />
                                                        </Tooltip>
                                                    }
                                                    {
                                                        /* Se delega la Attendance y el voto a la misma persona */
                                                        (
                                                            delegationPerson.VoteDelegation &&
                                                            delegationPerson.DelegationUserPrincipalName !== "" &&
                                                            delegationPerson.DelegationUserVotePrincipalName !== "" &&
                                                            delegationPerson.DelegationUserPrincipalName === delegationPerson.DelegationUserVotePrincipalName
                                                        ) &&
                                                        <>
                                                            <Tooltip content={strings.AttendanceDelegation} relationship='label' withArrow>
                                                                <Person20Regular className={styles.delegatedVote} />
                                                            </Tooltip>
                                                            <Tooltip content={strings.VoteDelegation} relationship='label' withArrow>
                                                                <HandLeft20Regular className={styles.delegatedVote} />
                                                            </Tooltip>
                                                        </>
                                                    }
                                                    {
                                                        /* Se delega la Attendance y el voto a distintas personas */
                                                        (
                                                            delegationPerson.VoteDelegation &&
                                                            delegationPerson.DelegationUserPrincipalName !== "" &&
                                                            delegationPerson.DelegationUserVotePrincipalName !== "" &&
                                                            delegationPerson.DelegationUserPrincipalName !== delegationPerson.DelegationUserVotePrincipalName
                                                        ) &&
                                                        <Tooltip content={strings.AttendanceDelegation} relationship='label' withArrow>
                                                            <Person20Regular className={styles.delegatedVote} />
                                                        </Tooltip>
                                                    }
                                                </div>
                                            </div>
                                        </div>
                                        :
                                        <div className={styles.meetingAssistants_alternativeSecondaryContainer}>
                                            <div>
                                                <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                                    <ArrowEnterLeft20Regular className={styles.meetingAssistants_diagonalArrow} style={{ transform: 'scaleX(-1)' }} />
                                                    <span className={styles.delegatedPending}> {strings.UserPendingApprodval + "..."} </span>
                                                </div>
                                            </div>
                                        </div>
                                    }
                                    {
                                        /*
                                            Si se delega la Attendance y el voto a distintas personas,
                                            se añade una fila para mostrar la persona a la que se le delega el voto
                                        */
                                        (
                                            delegationPerson.VoteDelegation &&
                                            delegationPerson.DelegationUserPrincipalName !== "" &&
                                            delegationPerson.DelegationUserVotePrincipalName !== "" &&
                                            delegationPerson.DelegationUserPrincipalName !== delegationPerson.DelegationUserVotePrincipalName
                                        ) &&
                                        <>
                                            <div className={styles.meetingAssistants_alternativeSecondaryContainerAux}                                            >
                                                <div>
                                                    <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                                        <Checkbox
                                                            checked={delegationPerson.HasAttended}
                                                            disabled={!this.state.editing}
                                                            onChange={(ev, data) => this.changeAttend(delegationPerson.UserPrincipalName, data.checked as boolean)}
                                                            style={{ visibility: 'hidden' }}
                                                        />
                                                        <Person
                                                            userId={delegationPerson.UserPrincipalName}
                                                            view={ViewType.threelines}
                                                            line3property='officeLocation'
                                                            className={styles.persona}
                                                            fetchImage={false}
                                                        />
                                                        {!locked && !editing && (
                                                            loadingAttendance === delegationPerson?.UserPrincipalName ?
                                                                <Spinner className={styles.deleteButton} size='tiny' />
                                                                :
                                                                this.renderDeleteBtn(delegationPerson))
                                                        }
                                                        <ArrowRight20Regular className={styles.meetingAssistants_rightArrow} />
                                                    </div>
                                                </div>
                                            </div>
                                            <div className={styles.meetingAssistants_alternativeSecondaryContainer}>
                                                <div>
                                                    <div className={styles.meetingAssistants_peopleRow_firstRow}>
                                                        <ArrowEnterLeft20Regular className={styles.meetingAssistants_diagonalArrow} style={{ transform: 'scaleX(-1)' }} />
                                                        <Person
                                                            userId={delegationPerson.DelegationUserVotePrincipalName}
                                                            view={ViewType.threelines}
                                                            line3property='officeLocation'
                                                            className={styles.persona}
                                                            fetchImage={false}
                                                        />
                                                        {
                                                            delegationPerson.VoteDelegation &&
                                                            <Tooltip content={strings.VoteDelegation} relationship='label' withArrow>
                                                                <HandLeft20Regular className={styles.delegatedVote} />
                                                            </Tooltip>
                                                        }
                                                    </div>
                                                </div>
                                            </div>
                                        </>
                                    }
                                </div>
                            );
                        })}
                    </div>
                </div>
            );
    }

}