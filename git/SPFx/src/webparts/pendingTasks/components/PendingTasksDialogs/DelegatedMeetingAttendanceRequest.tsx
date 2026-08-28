/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */
import * as React from 'react';
import * as strings from 'PendingTasksWebPartStrings';
import styles from '../PendingTasks.module.scss';
import {
    IDelegatedMeetingAttendanceRequest,
    IDelegatedMeetingAttendanceRequestprops,
    IDelegatedMeetingAttendanceRequestState
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
    Spinner,
    Textarea
} from '@fluentui/react-components';
import { CalendarReply20Filled } from '@fluentui/react-icons';
import { IUserTaskModel } from '../../../../service/BackendServiceModels/UserTaskModel';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import { ViewType } from '@microsoft/mgt-spfx';
import format from 'date-fns/format';
import ca from 'date-fns/locale/ca';
import es from 'date-fns/locale/es';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const TermIdOfTheAcceptOption: string = "9628620a-c31b-4903-9bb2-b305d51821bd";
const TermIdOfTheRefuseOption: string = "aac5817d-310f-4bfa-838a-1a1dda983bfb";

export default class DelegatedMeetingAttendanceRequest extends React.Component<IDelegatedMeetingAttendanceRequestprops, IDelegatedMeetingAttendanceRequestState> {

    constructor(props: IDelegatedMeetingAttendanceRequestprops) {
        super(props);
        this.state = {
            formInformation: undefined,
            meetingAttendanceDropdownErrorMessage: "",
            commentsErrorMessage: "",
            backendErrorMessage: "",
            sendingTheFormInformation: false,
            sendButtonIsDisabled: false,
            currentTask: this.props.pendingTask,
            isLoadingDetail: false
        };
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
                await bkService.updateTaskDelegate(
                    currentTask.BodyId,
                    currentTask.TaskId,
                    formInformation?.selectedTermIdFromTheMeetingAttendanceDropdown ?? "",
                    formInformation?.additionalComments ?? ""
                );

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
            formInformation
        } = this.state;

        let theFormInformationIsCorrect: boolean = true;

        if (
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown !== strings.Accept) &&
            (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown !== strings.Refuse)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                meetingAttendanceDropdownErrorMessage: strings.DropdownErrorMessage
            });
        }
        if (
            (formInformation?.additionalComments) &&
            (formInformation.additionalComments.length > 255)
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
            formInformation: undefined,
            meetingAttendanceDropdownErrorMessage: "",
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
        }).catch(error => {
            console.log(error);
            this.setState({
                currentTask: pendingTask,
                isLoadingDetail: false
            });
        });
    }

    public render(): React.ReactElement<IDelegatedMeetingAttendanceRequestprops> {
        const {
            locale,
            getThePendingTasksFromThisSite,
            DepartmentName
        } = this.props;
        const {
            formInformation,
            meetingAttendanceDropdownErrorMessage,
            commentsErrorMessage,
            backendErrorMessage,
            sendingTheFormInformation,
            sendButtonIsDisabled,
            currentTask,
            isLoadingDetail
        } = this.state;

        const locales = { "es-ES": es, "ca-ES": ca, "eu-ES": eu, "gl-ES": gl };

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
                                    <CalendarReply20Filled className={styles.icon} />
                                </div>
                                <div className={styles.cardText}>
                                    <span className={styles.cardTitle} >
                                        {strings.prefixDelegatedMeetingAttendanceRequest + ": " + currentTask.TaskTitle}
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
                                    {strings.prefixDelegatedMeetingAttendanceRequest + ": " + currentTask.TaskTitle}
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
                                            {/* Información de la delegación */}
                                            {
                                                /* Se delega la Attendance pero no el voto */
                                                (currentTask.DelegatedUserVote?.trim() === "") &&
                                                <div className={styles.messageDelegated}>
                                                    <Person userId={currentTask.RequestedBy} view={ViewType.oneline} />
                                                    <span className={styles.messageDelegatedText}> {strings.DelegateAssistanceTo} </span>
                                                    <Person userId={currentTask.DelegatedUser} view={ViewType.oneline} />
                                                    {((currentTask.Comment !== undefined) && (currentTask.Comment.trim() !== "")) && (
                                                        <span className={styles.messageDelegatedText}> {strings.Reason + ": " + currentTask.Comment} </span>
                                                    )}
                                                </div>
                                            }
                                            {
                                                /* Se delega la Attendance y el voto a la misma persona */
                                                (currentTask.DelegatedUser === currentTask.DelegatedUserVote) &&
                                                <div className={styles.messageDelegated}>
                                                    <Person userId={currentTask.RequestedBy} view={ViewType.oneline} />
                                                    <span className={styles.messageDelegatedText}> {strings.DelegateAssistanceAndVoteTo} </span>
                                                    <Person userId={currentTask.DelegatedUser} view={ViewType.oneline} />
                                                    {((currentTask.Comment !== undefined) && (currentTask.Comment.trim() !== "")) && (
                                                        <span className={styles.messageDelegatedText}> {strings.Reason + ": " + currentTask.Comment} </span>
                                                    )}
                                                </div>
                                            }
                                            {
                                                /* Se delega la Attendance y el voto a distintas personas */
                                                (
                                                    (currentTask.DelegatedUser !== currentTask.DelegatedUserVote) &&
                                                    (currentTask.DelegatedUserVote?.trim() !== "")
                                                ) &&
                                                <div className={styles.messageDelegated}>
                                                    <Person userId={currentTask.RequestedBy} view={ViewType.oneline} />
                                                    <span className={styles.messageDelegatedText}> {strings.DelegateAssistanceTo} </span>
                                                    <Person userId={currentTask.DelegatedUser} view={ViewType.oneline} />
                                                    <span className={styles.messageDelegatedText}> {strings.DelegateVoteTo} </span>
                                                    <Person userId={currentTask.DelegatedUserVote} view={ViewType.oneline} />
                                                    {((currentTask.Comment !== undefined) && (currentTask.Comment.trim() !== "")) && (
                                                        <span className={styles.messageDelegatedText}> {strings.Reason + ": " + currentTask.Comment} </span>
                                                    )}
                                                </div>
                                            }
                                            {/* Seleccionar una respuesta */}
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
                                                            const newFormInformation: IDelegatedMeetingAttendanceRequest = {
                                                                ...formInformation,
                                                                selectedOptionFromTheMeetingAttendanceDropdown: data.optionValue,
                                                                selectedTermIdFromTheMeetingAttendanceDropdown: (() => {
                                                                    if (data.optionValue === strings.Accept) {
                                                                        return TermIdOfTheAcceptOption;
                                                                    }
                                                                    else if (data.optionValue === strings.Refuse) {
                                                                        return TermIdOfTheRefuseOption;
                                                                    }
                                                                })()
                                                            };
                                                            this.setState({
                                                                formInformation: newFormInformation,
                                                                meetingAttendanceDropdownErrorMessage: ""
                                                            });
                                                        }
                                                    }}
                                                >
                                                    <Option>{strings.Accept}</Option>
                                                    <Option>{strings.Refuse}</Option>
                                                </Dropdown>
                                            </Field>
                                            {
                                                (meetingAttendanceDropdownErrorMessage && (meetingAttendanceDropdownErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {meetingAttendanceDropdownErrorMessage} </span>
                                                </Field>
                                            }
                                            {
                                                /* Comentarios adicionales */
                                                (formInformation?.selectedOptionFromTheMeetingAttendanceDropdown === strings.Refuse) &&
                                                <Field
                                                    style={{ marginBottom: "6px" }}
                                                    label={strings.AdditionalComments}
                                                >
                                                    <Textarea
                                                        placeholder={strings.WriteAComment}
                                                        onChange={(_ev, data): void => {
                                                            const additionalComments = data.value;
                                                            const commentsErrorMessage = (additionalComments && (additionalComments.length > 255)) ?
                                                                strings.CommentsErrorMessage : "";
                                                            this.setState({
                                                                formInformation: { ...formInformation, additionalComments },
                                                                commentsErrorMessage
                                                            });
                                                        }}
                                                        resize="vertical"
                                                    />
                                                </Field>
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