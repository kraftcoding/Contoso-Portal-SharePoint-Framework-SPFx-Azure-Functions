/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */
import * as React from 'react';
import * as strings from 'PendingTasksWebPartStrings';
import styles from '../PendingTasks.module.scss';
import {
    IMinutesApprodvalRequest,
    IMinutesApprodvalRequestprops,
    IMinutesApprodvalRequestState
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
    Textarea,
    Spinner
} from '@fluentui/react-components';
import { DocumentCheckmark20Filled } from '@fluentui/react-icons';
import { IUserTaskModel } from '../../../../service/BackendServiceModels/UserTaskModel';
import format from 'date-fns/format';
import ca from 'date-fns/locale/ca';
import es from 'date-fns/locale/es';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const TermIdOfTheAcceptOption: string = "9628620a-c31b-4903-9bb2-b305d51821bd";
const TermIdOfTheRefuseOption: string = "aac5817d-310f-4bfa-838a-1a1dda983bfb";
const TermIdOfTheModifyMinutesOption: string = "46707f88-14ef-456b-91f0-ae0785c89c7b";

export default class MinutesApprodvalRequest extends React.Component<IMinutesApprodvalRequestprops, IMinutesApprodvalRequestState> {

    constructor(props: IMinutesApprodvalRequestprops) {
        super(props);
        this.state = {
            formInformation: undefined,
            minutesApprodvalDropdownErrorMessage: "",
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
                await bkService.approdvalTaskMinutes(
                    currentTask.BodyId,
                    currentTask.TaskId,
                    formInformation?.selectedTermIdFromTheMinutesApprodvalDropdown ?? "",
                    formInformation?.additionalComments ?? ""
                );

                await getThePendingTasks();
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
            (formInformation?.selectedOptionFromTheMinutesApprodvalDropdown !== strings.Accept) &&
            (formInformation?.selectedOptionFromTheMinutesApprodvalDropdown !== strings.Refuse) &&
            (formInformation?.selectedOptionFromTheMinutesApprodvalDropdown !== strings.ModifyMinutes)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                minutesApprodvalDropdownErrorMessage: strings.DropdownErrorMessage
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
            minutesApprodvalDropdownErrorMessage: "",
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

    public render(): React.ReactElement<IMinutesApprodvalRequestprops> {
        const {
            locale,
            getThePendingTasksFromThisSite,
            DepartmentName
        } = this.props;
        const {
            formInformation,
            minutesApprodvalDropdownErrorMessage,
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
                                    <DocumentCheckmark20Filled className={styles.icon} />
                                </div>
                                <div className={styles.cardText}>
                                    <span className={styles.cardTitle}>
                                        {strings.prefixMinutesApprodvalRequest + ": " + currentTask.TaskTitle}
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
                                    {strings.prefixMinutesApprodvalRequest + ": " + currentTask.TaskTitle}
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
                                            {/* Seleccionar una respuesta */}
                                            <Field
                                                style={{ marginBottom: "6px" }}
                                                label={strings.MinutesApprodval}
                                            >
                                                <Dropdown
                                                    placeholder={strings.SelectAnOption}
                                                    aria-required={true}
                                                    value={formInformation?.selectedOptionFromTheMinutesApprodvalDropdown}
                                                    onOptionSelect={(_ev, data): void => {
                                                        if (data.optionValue) {
                                                            const newFormInformation: IMinutesApprodvalRequest = {
                                                                ...formInformation,
                                                                selectedOptionFromTheMinutesApprodvalDropdown: data.optionValue,
                                                                selectedTermIdFromTheMinutesApprodvalDropdown: (() => {
                                                                    if (data.optionValue === strings.Accept) {
                                                                        return TermIdOfTheAcceptOption;
                                                                    }
                                                                    else if (data.optionValue === strings.Refuse) {
                                                                        return TermIdOfTheRefuseOption;
                                                                    }
                                                                    else if (data.optionValue === strings.ModifyMinutes) {
                                                                        return TermIdOfTheModifyMinutesOption;
                                                                    }
                                                                })()
                                                            };
                                                            this.setState({
                                                                formInformation: newFormInformation,
                                                                minutesApprodvalDropdownErrorMessage: ""
                                                            });
                                                        }
                                                    }}
                                                >
                                                    <Option>{strings.Accept}</Option>
                                                    {/*<Option>{strings.Refuse}</Option>*/}
                                                    <Option>{strings.ModifyMinutes}</Option>
                                                </Dropdown>
                                            </Field>
                                            {
                                                (minutesApprodvalDropdownErrorMessage && (minutesApprodvalDropdownErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {minutesApprodvalDropdownErrorMessage} </span>
                                                </Field>
                                            }
                                            {
                                            /* Comentarios adicionales */
                                            (formInformation?.selectedOptionFromTheMinutesApprodvalDropdown === strings.ModifyMinutes) &&
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