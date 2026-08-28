/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */
import * as React from 'react';
import * as strings from 'PendingTasksWebPartStrings';
import styles from '../PendingTasks.module.scss';
import {
    IMeetingAttendanceRequest,
    IMeetingAttendanceRequestprops,
    IMeetingAttendanceRequestState
} from './IPendingTasksDialogs';
import {
    Dialog,
    DialogTrigger,
    DialogSurface,
    DialogBody,
    DialogTitle,
    DialogContent,
    DialogActions,
    Button,
    Field,
    Option,
    Dropdown,
    Checkbox,
    Textarea,
    Spinner,
    TextareaOnChangeData
} from '@fluentui/react-components';
import { Calendar20Filled } from '@fluentui/react-icons';
import { PeoplePicker } from '@microsoft/mgt-react/dist/es6/spfx';
import {
    IUserTaskDelegateModel,
    IUserTaskModel
} from '../../../../service/BackendServiceModels/UserTaskModel';
import {
    IDynamicPerson,
    PersonType,
    //UserType
} from '@microsoft/mgt-spfx';
import format from 'date-fns/format';
import ca from 'date-fns/locale/ca';
import es from 'date-fns/locale/es';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const TermIdOfTheAcceptOption: string = "9628620a-c31b-4903-9bb2-b305d51821bd";
const TermIdOfTheRefuseOption: string = "aac5817d-310f-4bfa-838a-1a1dda983bfb";
const TermIdOfTheDelegateOption: string = "";
const TermIdOfTheOnlineAndInPersonOption: string = "e49c8fe8-87f5-4b62-b6ce-0df240b7ec7f";
const TermIdOfTheOnlineOption: string = "8f4e0e22-2c3f-4454-8210-f4151dceeb89";
const TermIdOfTheInPersonOption: string = "3a16b292-5616-4ef9-8a56-2b36c20bb992";

export default class MeetingAttendanceRequest extends React.Component<IMeetingAttendanceRequestprops, IMeetingAttendanceRequestState> {

    constructor(props: IMeetingAttendanceRequestprops) {
        super(props);
        this.state = {
            formInformation: {
                personSelectedToDelegateTheAttendance: [],
                personSelectedToDelegateTheVote: []
            },
            meetingAttendanceDropdownErrorMessage: "",
            attendanceTypeDropdownErrorMessage: "",
            delegatedAttendancePeoplePickerErrorMessage: "",
            delegatedVotePeoplePickerErrorMessage: "",
            commentsErrorMessage: "",
            backendErrorMessage: "",
            sendingTheFormInformation: false,
            sendButtonIsDisabled: false,
            membersGroupId: "",
            asistenteGroupId: "",
            currentTask: this.props.pendingTask,
            isLoadingDetail: false
        };
    }

    public async componentDidMount(): promise<void> {
        const {
            spService,
            pendingTask
        } = this.props;

        const bodyId: string = pendingTask?.BodyId;
        try {
            const membersGroupId: string = await spService.getGroupId(`${bodyId}_members`);
            const asistenteGroupId: string = await spService.getGroupId(`${bodyId}_asistentemembers`);
            this.setState({ membersGroupId, asistenteGroupId });
        }
        catch (error) {
            console.error(error);
        }
    }

    private async submitTheFormInformation(): promise<void> {
        const {
            bkService,
            getThePendingTasks
        } = this.props;
        const {
            formInformation,
            currentTask
        } = this.state;

        const theFormInformationIsCorrect: boolean = this.checkTheFormInformation();

        if (theFormInformationIsCorrect) {
            this.setState({
                sendingTheFormInformation: true
            });
            try {
                if (
                    (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Accept) ||
                    (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Refuse)
                ) {
                    await bkService.updateTask(
                        currentTask.BodyId,
                        currentTask.TaskId,
                        formInformation.selectedTermIdFromTheMeetingAttendanceDropdown ?? "",
                        formInformation.selectedTermIdFromTheAttendanceTypeDropdown ?? currentTask.EventAttendanceTypeId,
                        formInformation.additionalComments ?? ""
                    );
                }
                else if (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) {
                    const delegatedMeetingAttendanceTask: IUserTaskDelegateModel = {
                        TaskTitle: currentTask.TaskTitle,
                        TaskStartDate: currentTask.TaskStartDate,
                        TaskEndDate: currentTask.TaskEndDate,
                        Comment: formInformation.additionalComments ?? "",
                        SharedEventId: currentTask.SharedEventId,
                        TaskParentId: currentTask.TaskId,
                        DelegateTo: (
                            formInformation.personSelectedToDelegateTheAttendance &&
                            (formInformation.personSelectedToDelegateTheAttendance.length > 0)
                        ) ? formInformation.personSelectedToDelegateTheAttendance[0].userPrincipalName : '',
                        DelegateVote: formInformation.voteIsDelegated ?? false,
                        DelegatedUserVote: (
                            formInformation.personSelectedToDelegateTheVote &&
                            (formInformation.personSelectedToDelegateTheVote.length > 0)
                        ) ? formInformation.personSelectedToDelegateTheVote[0].userPrincipalName : ''
                    }

                    await bkService.createDelegateTask(currentTask.BodyId, delegatedMeetingAttendanceTask);
                }

                await getThePendingTasks();

                this.setState({
                    sendingTheFormInformation: false
                });
            }
            catch (error) {
                console.error(error);

                this.setState({
                    backendErrorMessage: strings.BackendErrorMessage,
                    sendingTheFormInformation: false,
                    sendButtonIsDisabled: true
                });
            }
        }
    }

    private checkTheFormInformation(): boolean {
        const {
            context
        } = this.props;
        const {
            formInformation,
            currentTask
        } = this.state;

        let theFormInformationIsCorrect: boolean = true;
        const currentUserEmail: string = context.pageContext.user.email;

        /* Validación de Attendance a la Meeting*/
        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown !== strings.Accept) &&
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown !== strings.Refuse) &&
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown !== strings.Delegate)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                meetingAttendanceDropdownErrorMessage: strings.DropdownErrorMessage
            });
        }

        /* Validación de Tipo de Attendance */
        if (
            (currentTask.EventAttendanceTypeId === TermIdOfTheOnlineAndInPersonOption) &&
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Accept)
        ) {
            if (
                (formInformation?.selectedOptionFromTheAttendanceTypeDropdown !== strings.Online) &&
                (formInformation?.selectedOptionFromTheAttendanceTypeDropdown !== strings.InPerson)
            ) {
                theFormInformationIsCorrect = false;
                this.setState({
                    attendanceTypeDropdownErrorMessage: strings.DropdownErrorMessage
                });
            }
        }

        /* Validación de Delegar la Meeting */
        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) &&
            (formInformation?.personSelectedToDelegateTheAttendance?.length === 0)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                delegatedAttendancePeoplePickerErrorMessage: strings.PeoplePickerErrorMessage
            });
        }
        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) &&
            (formInformation?.personSelectedToDelegateTheVote?.length === 0) &&
            (formInformation?.voteIsDelegated === true)

        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                delegatedVotePeoplePickerErrorMessage: strings.PeoplePickerErrorMessage
            });
        }
        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) &&
            (formInformation?.personSelectedToDelegateTheAttendance) &&
            (formInformation?.personSelectedToDelegateTheAttendance[0]?.mail === currentUserEmail)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                delegatedAttendancePeoplePickerErrorMessage: strings.PeoplePickerDelegationErrorMessage
            });
        }
        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) &&
            (formInformation?.personSelectedToDelegateTheVote) &&
            (formInformation?.personSelectedToDelegateTheVote[0]?.mail === currentUserEmail)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                delegatedVotePeoplePickerErrorMessage: strings.PeoplePickerDelegationErrorMessage
            });
        }

        /* Validación de Comentarios adicionales */
        if (
            (formInformation?.additionalComments) &&
            (formInformation?.additionalComments.length > 255)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                commentsErrorMessage: strings.CommentsErrorMessage
            });
        }

        return theFormInformationIsCorrect;
    }

    private onClickTask(): void {
        const {
            bkService,
            pendingTask
        } = this.props;

        this.setState({
            formInformation: {
                personSelectedToDelegateTheAttendance: [],
                personSelectedToDelegateTheVote: []
            },
            meetingAttendanceDropdownErrorMessage: "",
            attendanceTypeDropdownErrorMessage: "",
            delegatedAttendancePeoplePickerErrorMessage: "",
            commentsErrorMessage: "",
            backendErrorMessage: "",
            sendingTheFormInformation: false,
            sendButtonIsDisabled: false,
            isLoadingDetail: true
        });

        bkService.getUserTaskByTaskId(pendingTask.TaskId, pendingTask.BodyId).then((res: IUserTaskModel): void => {
            this.setState({
                currentTask: res,
                isLoadingDetail: false
            });
        }
        ).catch(error => {
            console.log(error);
            this.setState({
                currentTask: pendingTask,
                isLoadingDetail: false
            });
        });
    }

    public render(): React.ReactElement<IMeetingAttendanceRequestprops> {
        const {
            locale,
            context,
            getThePendingTasksFromThisSite,
            DepartmentName
        } = this.props;
        const {
            formInformation,
            meetingAttendanceDropdownErrorMessage,
            attendanceTypeDropdownErrorMessage,
            delegatedAttendancePeoplePickerErrorMessage,
            delegatedVotePeoplePickerErrorMessage,
            commentsErrorMessage,
            backendErrorMessage,
            sendingTheFormInformation,
            sendButtonIsDisabled,
            membersGroupId,
            asistenteGroupId,
            currentTask,
            isLoadingDetail
        } = this.state;

        const locales = { "es-ES": es, "ca-ES": ca, "eu-ES": eu, "gl-ES": gl };
        const currentUserEmail: string = context.pageContext.user.email;

        return (
            <>
                {/* @ts-ignore */}
                <Dialog>
                    <DialogTrigger disableButtonEnhancement>
                        <Button
                            className={styles.cardButton}
                            onClick={() => this.onClickTask()}
                        >
                            {/* Contenido del botón */}
                            <div className={styles.card}>
                                <div className={styles.cardIcon}>
                                    <Calendar20Filled className={styles.icon} />
                                </div>
                                <div className={styles.cardText}>
                                    <span className={styles.cardTitle}>
                                        {strings.prefixMeetingAttendanceRequest + ": " + currentTask.TaskTitle}
                                    </span>
                                    {
                                        (!getThePendingTasksFromThisSite) &&
                                        <span className={styles.cardOrgan}> {DepartmentName} </span>
                                    }
                                    {
                                        (currentTask.TaskEndDate && new Date(currentTask.TaskEndDate).getTime()) &&
                                        <span className={styles.cardDate}>
                                            {/* @ts-ignore */}
                                            {format(new Date(currentTask.TaskEndDate), 'PPPP', { locale: locales[locale] }).replace(/^\w/, (character: string): string => character.toUpperCase())}
                                        </span>
                                    }
                                </div>
                            </div>
                        </Button>
                    </DialogTrigger>
                    <DialogSurface>
                        <DialogBody>
                            <DialogTitle
                                style={{
                                    fontSize: "16px",
                                    lineHeight: "22px",
                                    overflow: "hidden",
                                    display: "-webkit-box",
                                    WebkitBoxOrient: "vertical",
                                    WebkitLineClamp: 2
                                }}
                            >
                                <span className={styles.cardTitle}>
                                    {strings.prefixMeetingAttendanceRequest + ": " + currentTask.TaskTitle}
                                </span>
                            </DialogTitle>
                            <DialogContent>
                                {
                                    (isLoadingDetail) ?
                                        <Spinner className={styles.spinnerLoadingModal}></Spinner>
                                        :
                                        <>
                                            {/* Nombre del órgano */}
                                            {
                                                (!getThePendingTasksFromThisSite) &&
                                                <Field
                                                    style={{
                                                        marginBottom: "6px",
                                                        fontSize: "14px",
                                                        fontWeight: "400",
                                                        overflow: "hidden",
                                                        display: "-webkit-box",
                                                        WebkitBoxOrient: "vertical",
                                                        WebkitLineClamp: 1
                                                    }}
                                                >
                                                    <span> {DepartmentName} </span>
                                                </Field>
                                            }
                                            {/* Attendance a la Meeting */}
                                            <Field
                                                style={{ marginBottom: "6px" }}
                                                label={strings.MeetingAttendance}
                                            >
                                                <Dropdown
                                                    placeholder={strings.SelectAnOption}
                                                    aria-required={true}
                                                    value={formInformation?.selectedOptionFromTheMeetingAttendanceDropdown}
                                                    onOptionSelect={(_ev, data): void => {
                                                        if (data.optionValue) {
                                                            const newFormInformation: IMeetingAttendanceRequest = {
                                                                ...formInformation,
                                                                personSelectedToDelegateTheAttendance: [],
                                                                selectedOptionFromTheMeetingAttendanceDropdown: data.optionValue,
                                                                selectedTermIdFromTheMeetingAttendanceDropdown: (() => {
                                                                    if (data.optionValue === strings.Accept) {
                                                                        return TermIdOfTheAcceptOption;
                                                                    }
                                                                    else if (data.optionValue === strings.Refuse) {
                                                                        return TermIdOfTheRefuseOption;
                                                                    }
                                                                    else if (data.optionValue === strings.Delegate) {
                                                                        return TermIdOfTheDelegateOption;
                                                                    }
                                                                })()
                                                            };
                                                            if (data.optionValue === strings.Delegate) {
                                                                newFormInformation.attendanceIsDelegated = true;
                                                                newFormInformation.voteIsDelegated = currentTask.Vote;
                                                                newFormInformation.personSelectedToDelegateTheAttendance = [];
                                                                newFormInformation.personSelectedToDelegateTheVote = [];
                                                            }
                                                            else {
                                                                newFormInformation.attendanceIsDelegated = false;
                                                                newFormInformation.voteIsDelegated = false;
                                                                newFormInformation.personSelectedToDelegateTheAttendance = [];
                                                                newFormInformation.personSelectedToDelegateTheVote = [];
                                                            }
                                                            this.setState({
                                                                formInformation: newFormInformation,
                                                                meetingAttendanceDropdownErrorMessage: "",
                                                                delegatedAttendancePeoplePickerErrorMessage: "",
                                                                delegatedVotePeoplePickerErrorMessage: ""
                                                            });
                                                        }
                                                    }}
                                                >
                                                    <Option>{strings.Accept}</Option>
                                                    <Option>{strings.Refuse}</Option>
                                                    {
                                                        (currentTask?.Vote) &&
                                                        <Option>{strings.Delegate}</Option>
                                                    }
                                                </Dropdown>
                                            </Field>
                                            {
                                                (meetingAttendanceDropdownErrorMessage && (meetingAttendanceDropdownErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {meetingAttendanceDropdownErrorMessage} </span>
                                                </Field>
                                            }
                                            {/* Tipo de Attendance */}
                                            {
                                                (
                                                    (currentTask.EventAttendanceTypeId === TermIdOfTheOnlineAndInPersonOption) &&
                                                    (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Accept)
                                                ) &&
                                                <Field
                                                    style={{ marginBottom: "6px" }}
                                                    label={strings.AttendanceType}
                                                >
                                                    <Dropdown
                                                        placeholder={strings.SelectAnOption}
                                                        aria-required={true}
                                                        value={formInformation?.selectedOptionFromTheAttendanceTypeDropdown}
                                                        onOptionSelect={(_ev, data): void => {
                                                            if (data.optionValue) {
                                                                const newFormInformation: IMeetingAttendanceRequest = {
                                                                    ...formInformation,
                                                                    selectedOptionFromTheAttendanceTypeDropdown: data.optionValue,
                                                                    selectedTermIdFromTheAttendanceTypeDropdown: (() => {
                                                                        if (data.optionValue === strings.Online) {
                                                                            return TermIdOfTheOnlineOption;
                                                                        }
                                                                        else if (data.optionValue === strings.InPerson) {
                                                                            return TermIdOfTheInPersonOption;
                                                                        }
                                                                    })()
                                                                };
                                                                this.setState({
                                                                    formInformation: newFormInformation,
                                                                    attendanceTypeDropdownErrorMessage: ""
                                                                });
                                                            }
                                                        }}
                                                    >
                                                        <Option>{strings.Online}</Option>
                                                        <Option>{strings.InPerson}</Option>
                                                    </Dropdown>
                                                </Field>
                                            }
                                            {
                                                (attendanceTypeDropdownErrorMessage && (attendanceTypeDropdownErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {attendanceTypeDropdownErrorMessage} </span>
                                                </Field>
                                            }
                                            {/* Delegar la Attendance */}
                                            {
                                                (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Delegate) &&
                                                <>
                                                    <Field
                                                        style={{ marginBottom: "6px" }}
                                                        label={strings.DelegateTheMeeting}
                                                    >
                                                        <div>
                                                            <Checkbox
                                                                label={strings.Attendance}
                                                                checked={formInformation?.attendanceIsDelegated}
                                                                disabled={true}
                                                            />
                                                            <PeoplePicker
                                                                type={PersonType.person}
                                                                //userType={UserType.user}
                                                                placeholder={strings.SelectAPerson}
                                                                selectionMode="single"
                                                                groupIds={[membersGroupId, asistenteGroupId]}
                                                                selectedPeople={formInformation.personSelectedToDelegateTheAttendance || []}
                                                                selectionChanged={async (persons: CustomEvent<IDynamicPerson[]>): promise<void> => {
                                                                    if (persons.detail.length > 0) {
                                                                        persons.detail.forEach((person: IDynamicPerson): void => {
                                                                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                                                            const selectedUser: any = person;
                                                                            const selectedUserEmail: string = selectedUser.mail;
                                                                            if (selectedUserEmail === currentUserEmail) {
                                                                                this.setState({
                                                                                    formInformation: {
                                                                                        ...formInformation,
                                                                                        personSelectedToDelegateTheAttendance: [person]
                                                                                    },
                                                                                    delegatedAttendancePeoplePickerErrorMessage: strings.PeoplePickerDelegationErrorMessage
                                                                                });
                                                                            }
                                                                            else {
                                                                                this.setState({
                                                                                    formInformation: {
                                                                                        ...formInformation,
                                                                                        personSelectedToDelegateTheAttendance: [person]
                                                                                    },
                                                                                    delegatedAttendancePeoplePickerErrorMessage: ""
                                                                                });
                                                                            }
                                                                        });
                                                                    }
                                                                    else {
                                                                        this.setState({
                                                                            formInformation: {
                                                                                ...formInformation,
                                                                                personSelectedToDelegateTheAttendance: []
                                                                            },
                                                                            delegatedAttendancePeoplePickerErrorMessage: strings.PeoplePickerErrorMessage
                                                                        });
                                                                    }
                                                                }}
                                                            />
                                                        </div>
                                                        {
                                                            (delegatedAttendancePeoplePickerErrorMessage && (delegatedAttendancePeoplePickerErrorMessage !== "")) &&
                                                            <Field style={{ color: "red", marginTop: "6px" }}>
                                                                <span> {delegatedAttendancePeoplePickerErrorMessage} </span>
                                                            </Field>
                                                        }
                                                        {
                                                            (currentTask?.Vote) &&
                                                            <div>
                                                                <Checkbox
                                                                    label={strings.Vote}
                                                                    checked={formInformation?.voteIsDelegated}
                                                                    onChange={(_ev, data): void =>
                                                                        this.setState({
                                                                            formInformation: {
                                                                                ...formInformation,
                                                                                voteIsDelegated: !!data.checked,
                                                                                personSelectedToDelegateTheVote: []
                                                                            },
                                                                            delegatedVotePeoplePickerErrorMessage: ""
                                                                        })
                                                                    }
                                                                />
                                                                <PeoplePicker
                                                                    disabled={!formInformation?.voteIsDelegated}
                                                                    type={PersonType.person}
                                                                    //userType={UserType.user}
                                                                    placeholder={strings.SelectAPerson}
                                                                    selectionMode="single"
                                                                    groupIds={[membersGroupId, asistenteGroupId]}
                                                                    selectedPeople={formInformation.personSelectedToDelegateTheVote || []}
                                                                    selectionChanged={async (persons: CustomEvent<IDynamicPerson[]>): promise<void> => {
                                                                        if (persons.detail.length > 0) {
                                                                            persons.detail.forEach((person: IDynamicPerson): void => {
                                                                                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                                                                                const selectedUser: any = person;
                                                                                const selectedUserEmail: string = selectedUser.mail;
                                                                                if (selectedUserEmail === currentUserEmail) {
                                                                                    this.setState({
                                                                                        formInformation: {
                                                                                            ...formInformation,
                                                                                            personSelectedToDelegateTheVote: [person]
                                                                                        },
                                                                                        delegatedVotePeoplePickerErrorMessage: strings.PeoplePickerDelegationErrorMessage
                                                                                    });
                                                                                }
                                                                                else {
                                                                                    this.setState({
                                                                                        formInformation: {
                                                                                            ...formInformation,
                                                                                            personSelectedToDelegateTheVote: [person]
                                                                                        },
                                                                                        delegatedVotePeoplePickerErrorMessage: ""
                                                                                    });
                                                                                }
                                                                            });
                                                                        }
                                                                        else {
                                                                            this.setState({
                                                                                formInformation: {
                                                                                    ...formInformation,
                                                                                    personSelectedToDelegateTheVote: []
                                                                                },
                                                                                delegatedVotePeoplePickerErrorMessage: strings.PeoplePickerErrorMessage
                                                                            });
                                                                        }
                                                                    }}
                                                                />
                                                            </div>}
                                                        {
                                                            (delegatedVotePeoplePickerErrorMessage && (delegatedVotePeoplePickerErrorMessage !== "")) &&
                                                            <Field style={{ color: "red", marginTop: "6px" }}>
                                                                <span> {delegatedVotePeoplePickerErrorMessage} </span>
                                                            </Field>
                                                        }
                                                    </Field>
                                                    {/* Comentarios adicionales */}
                                                    <Field
                                                        style={{ marginBottom: "6px" }}
                                                        label={strings.AdditionalComments}
                                                    >
                                                        <Textarea
                                                            placeholder={strings.WriteAComment}
                                                            onChange={(_ev, data: TextareaOnChangeData): void => {
                                                                const additionalComments = data.value;
                                                                const commentsErrorMessage = (additionalComments && (additionalComments.length > 255)) ?
                                                                    strings.CommentsErrorMessage : "";
                                                                this.setState({
                                                                    formInformation: { ...formInformation as IMeetingAttendanceRequest, additionalComments },
                                                                    commentsErrorMessage
                                                                });
                                                            }}
                                                            resize="vertical"
                                                        />
                                                    </Field>
                                                </>

                                            }

                                            {
                                                (commentsErrorMessage && (commentsErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {commentsErrorMessage} </span>
                                                </Field>
                                            }
                                            {
                                                (backendErrorMessage && (backendErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span > {backendErrorMessage} </span>
                                                </Field>
                                            }
                                        </>
                                }
                            </DialogContent>
                            <DialogActions>
                                {
                                    (sendingTheFormInformation) &&
                                    <Spinner size='tiny' label={strings.Sending + "..."} />
                                }
                                <DialogTrigger disableButtonEnhancement>
                                    <Button appearance="secondary">
                                        {strings.RuleOut}
                                    </Button>
                                </DialogTrigger>
                                <Button
                                    appearance="primary"
                                    disabled={sendingTheFormInformation || sendButtonIsDisabled || isLoadingDetail}
                                    onClick={(): promise<void> => this.submitTheFormInformation()}
                                >
                                    {strings.Send}
                                </Button>
                            </DialogActions>
                        </DialogBody>
                    </DialogSurface>
                </Dialog >
            </>
        );
    }

}