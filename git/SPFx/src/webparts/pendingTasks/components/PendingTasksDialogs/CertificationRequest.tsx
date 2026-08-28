/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable no-void */
import * as React from 'react';
import * as strings from 'PendingTasksWebPartStrings';
import styles from '../PendingTasks.module.scss';
import {
    ICertificationRequest,
    ICertificationRequestprops,
    ICertificationRequestState
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
    Spinner
} from '@fluentui/react-components';
import { DocumentRibbon20Filled } from '@fluentui/react-icons';
import { IUserTaskModel } from '../../../../service/BackendServiceModels/UserTaskModel';
import format from 'date-fns/format';
import ca from 'date-fns/locale/ca';
import es from 'date-fns/locale/es';
import eu from 'date-fns/locale/eu';
import gl from 'date-fns/locale/gl';

const TermIdOfTheAcceptOption: string = "9628620a-c31b-4903-9bb2-b305d51821bd";
const TermIdOfTheRefuseOption: string = "aac5817d-310f-4bfa-838a-1a1dda983bfb";

export default class CertificationRequest extends React.Component<ICertificationRequestprops, ICertificationRequestState> {

    constructor(props: ICertificationRequestprops) {
        super(props);
        this.state = {
            formInformation: undefined,
            certificationApprodvalDropdownErrorMessage: "",
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
                await bkService.approdvalTaskCertification(
                    currentTask.BodyId,
                    currentTask.TaskId,
                    formInformation?.selectedTermIdFromTheCertificationApprodvalDropdown ?? ""
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
            (formInformation?.selectedOptionFromTheCertificationApprodvalDropdown !== strings.Accept) &&
            (formInformation?.selectedOptionFromTheCertificationApprodvalDropdown !== strings.Refuse)
        ) {
            theFormInformationIsCorrect = false;
            this.setState({
                certificationApprodvalDropdownErrorMessage: strings.DropdownErrorMessage
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
            isLoadingDetail: true,
            formInformation: undefined,
            certificationApprodvalDropdownErrorMessage: "",
            backendErrorMessage: "",
            sendingTheFormInformation: false,
            sendButtonIsDisabled: false
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

    public render(): React.ReactElement<ICertificationRequestprops> {
        const {
            locale,
            getThePendingTasksFromThisSite,
            DepartmentName
        } = this.props;
        const {
            formInformation,
            certificationApprodvalDropdownErrorMessage,
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
                                    <DocumentRibbon20Filled className={styles.icon} />
                                </div>
                                <div className={styles.cardText}>
                                    <span className={styles.cardTitle}>
                                        {strings.prefixCertificationRequest + ": " + currentTask.TaskTitle}
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
                                    {strings.prefixCertificationRequest + ": " + currentTask.TaskTitle}
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
                                            {
                                                /* Información de la Request de certificación */
                                                (currentTask.AgreementTitle && (currentTask.AgreementTitle.trim() !== "")) &&
                                                <div className={styles.messageDelegated}>
                                                    <span> {strings.ThereIsACertificationRequest + ": " + currentTask.AgreementTitle} </span>
                                                </div>
                                            }
                                            {/* Seleccionar una respuesta */}
                                            <Field
                                                style={{ marginBottom: "6px" }}
                                                label={strings.CertificationApprodval}
                                            >
                                                <Dropdown
                                                    placeholder={strings.SelectAnOption}
                                                    aria-required={true}
                                                    value={formInformation?.selectedOptionFromTheCertificationApprodvalDropdown}
                                                    onOptionSelect={(_ev, data): void => {
                                                        if (data.optionValue) {
                                                            const newFormInformation: ICertificationRequest = {
                                                                ...formInformation,
                                                                selectedOptionFromTheCertificationApprodvalDropdown: data.optionValue,
                                                                selectedTermIdFromTheCertificationApprodvalDropdown: (() => {
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
                                                                certificationApprodvalDropdownErrorMessage: ""
                                                            });
                                                        }
                                                    }}
                                                >
                                                    <Option>{strings.Accept}</Option>
                                                    <Option>{strings.Refuse}</Option>
                                                </Dropdown>
                                            </Field>
                                            {
                                                (certificationApprodvalDropdownErrorMessage && (certificationApprodvalDropdownErrorMessage !== "")) &&
                                                <Field style={{ color: "red", marginBottom: "6px" }}>
                                                    <span> {certificationApprodvalDropdownErrorMessage} </span>
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