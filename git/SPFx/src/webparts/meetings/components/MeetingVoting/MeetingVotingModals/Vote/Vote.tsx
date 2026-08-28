import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './Vote.module.scss';
import {
    IVoteprops,
    IVoteState
} from '../IMeetingVotingModals';
import {
    ButtonActions,
    VotingContext,
    VotingDialogActions,
    VotingTermsIds
} from '../../IMeetingVoting';
import {
    CheckmarkRegular,
    DismissCircleRegular,
    DismissFilled,
    ErrorCircleRegular,
    HandRightRegular,
    VoteRegular
} from '@fluentui/react-icons';
import {
    Button,
    Dialog,
    DialogActions,
    DialogBody,
    DialogSurface,
    DialogTrigger,
    Radio,
    RadioGroup,
    Spinner
} from '@fluentui/react-components';
import { Term } from '../../../../../../models/ITag';
import { Logger } from '../../../../../../utils/Logger';
import { IsNullOrECNTy } from '../../../../../../utils/Utils';
import {
    AgreementStatus,
    IVote
} from '../../../../../../service/BackendServiceModels/EventModels';

export default class Vote extends React.Component<IVoteprops, IVoteState> {

    static contextType = VotingContext;
    context!: React.ContextType<typeof VotingContext>;
    private testvVote: string = "";

    private voteTaxonomyCodes: Record<string, Term> = {};

    constructor(props: IVoteprops) {
        super(props);
        this.state = {
            selectedButton: "",
            secondScreen: false,
            errorMsg: false,
            isLoading: false,
            selectedVotingId: ""
        }
    }

    public async componentDidMount(): promise<void> {
        const { selectedVotingId } = this.props;

        this.setState({ selectedVotingId });

        this.setState({
            selectedButton: "",
            errorMsg: false,
            secondScreen: false
        });
    }

    public async componentDidUpdate(testvprops: Readonly<IVoteprops>): promise<void> {
        const { selectedVotingId } = this.props;

        if (testvprops.selectedVotingId !== selectedVotingId) {
            this.setState({ selectedVotingId });
        }

        if (IsNullOrECNTy(this.voteTaxonomyCodes)) {
            this.getStorageCodes();
        }
    }

    private getStorageCodes(): void {
        const { termStorage } = this.context;

        if (!IsNullOrECNTy(termStorage)) {
            this.voteTaxonomyCodes = termStorage?.[VotingTermsIds.VoteOptionsSet];
        }
    }

    private closeVoteModal = async (): promise<void> => {
        // TODO: Revisar el cierre del modal
        this.context.onDismissModal();

        this.testvVote = "";

        this.setState({
            selectedButton: "",
            errorMsg: false,
            secondScreen: false,
        });
    }

    private saveFirstStepVoteModal = (): void => {
        const { selectedButton } = this.state;

        if ((selectedButton === "") || (selectedButton === "-1")) {
            this.setState({
                errorMsg: true
            });
        }
        else {
            this.setState({
                secondScreen: true,
                errorMsg: false
            });
        }
    }

    private returnToFirstScreenVoteModal = (): void => {
        this.setState({
            secondScreen: false
        });
    }

    private saveAndCloseVoteModal = async (): promise<void> => {
        const { refresh } = this.props;
        const { selectedButton } = this.state;
        const {
            votingSelected,
            bkService,
            event
        } = this.context;

        // TODO: Revisar el cierre del modal

        if (votingSelected) {
            this.setState({ isLoading: true });
            try {
                const vote: IVote = {
                    Id: votingSelected.Id,
                    VoteId: selectedButton,
                };
                await bkService.updateVoteByUser(event.BodyId, event.Id, vote);
            }
            catch (error) {
                Logger.error('Failed to update Vote', error, this.context.context);
                throw error;
            }
        }

        await new promise<void>((resolve) => {
            this.context.onDismissModal();
            resolve();
        });

        await refresh();

        this.testvVote = "";

        this.setState({
            selectedButton: "",
            secondScreen: false,
            isLoading: false
        });
    }

    public render(): React.ReactElement<IVoteprops> {
        const {
            selectedButton,
            secondScreen,
            errorMsg,
            isLoading,
            selectedVotingId
        } = this.state;
        const {
            event,
            context,
            votingSelected,
            votingDialog,
            votingDetailMap,
            votingIdentity
        } = this.context;

        const DepartmentName: string = (context?.pageContext.web.title) ?? "";
        const canVote: boolean = (votingIdentity !== "") && (votingIdentity !== "NOVOTE");
        const votingFinishedAndUserDidNotVote: boolean = (
            (
                (votingSelected?.StatusId === AgreementStatus.Approdved) ||
                (votingSelected?.StatusId === AgreementStatus.UnanimouslyApprodved) ||
                (votingSelected?.StatusId === AgreementStatus.Declined) ||
                (votingSelected?.StatusId === AgreementStatus.OverTheTable)
            ) &&
            (
                (votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] === undefined) ||
                (votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] === null)
            )
        );

        if (votingSelected?.Id && canVote) {
            if (
                (votingDialog === VotingDialogActions.Vote) &&
                (selectedButton === "") &&
                (votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] !== undefined) &&
                (votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id] !== null)
            ) {
                this.testvVote = votingDetailMap?.[votingIdentity]?.Votes?.[votingSelected?.Id]

                this.setState({
                    selectedButton: this.testvVote,
                    secondScreen: (this.testvVote !== "") ? true : false
                });
            }
        }

        if ((votingDialog === VotingDialogActions.Vote) || (votingSelected?.Id === selectedVotingId)) {
            return (
                <Dialog open={(votingDialog === VotingDialogActions.Vote) || (votingSelected?.Id === selectedVotingId)}>
                    <DialogSurface className={styles.vote}>
                        {/* Cabecera */}
                        <div className={styles.voteHeader}>
                            {/* Icono */}
                            <div className={styles.voteLogo}>
                                <VoteRegular />
                            </div>
                            {/* Título del órgano */}
                            <div className={styles.voteTitle}>
                                {DepartmentName}
                            </div>
                            {
                                /* Botón de cerrar */
                                (!secondScreen || secondScreen || votingFinishedAndUserDidNotVote) &&
                                <Button
                                    icon={<DismissFilled />}
                                    appearance="transparent"
                                    className={styles.voteDismissIcon}
                                    onClick={this.closeVoteModal}
                                    disabled={isLoading}
                                />
                            }
                        </div>
                        {/* Subcabecera */}
                        <div className={styles.voteSubheader}>
                            {/* Título de la Meeting */}
                            <div className={styles.voteFirstText}>
                                {event.Title}
                            </div>
                            {/* Título de la votación */}
                            <div className={styles.voteSecondText}>
                                {
                                    (votingSelected?.Order && votingSelected?.Title) ?
                                        (votingSelected.Order + "- " + votingSelected.Title)
                                        :
                                        undefined
                                }
                            </div>
                        </div>
                        {
                            /* Enviando el voto */
                            (isLoading) &&
                            <div className={styles.voteSpinner}>
                                <Spinner /> <span> {strings.SendingVote + "..."} </span>
                            </div>
                        }
                        {
                            /* Primera pantalla - Selección de voto */
                            (!secondScreen && canVote && !isLoading && !votingFinishedAndUserDidNotVote) &&
                            <DialogBody className={styles.voteBody}>
                                {/* Opciones para votar */}
                                <RadioGroup className={styles.voteOptions} defaultValue={selectedButton}>
                                    {/* A favor */}
                                    <Button
                                        className={`
                                        ${styles.voteSingleButton} 
                                        ${styles.voteSingleButtonGreen} 
                                        ${(selectedButton === ButtonActions.AFavor) ?
                                                styles.selectedButton
                                                :
                                                styles.dimmedButton
                                            }
                                    `}
                                        onClick={(): void => this.setState({
                                            selectedButton: ButtonActions.AFavor,
                                            errorMsg: false
                                        })}
                                    >
                                        <Radio
                                            value={ButtonActions.AFavor}
                                            checked={selectedButton === ButtonActions.AFavor}
                                        />
                                        <span className={styles.voteSingleButtonText}>
                                            {strings.Approdve}
                                        </span>
                                        <CheckmarkRegular />
                                    </Button>
                                    {/* Abstención */}
                                    <Button
                                        className={`
                                        ${styles.voteSingleButton} 
                                        ${styles.voteSingleButtonGrey}                                   
                                        ${(selectedButton === ButtonActions.Abstencion) ?
                                                styles.selectedButton
                                                :
                                                styles.dimmedButton
                                            }
                                    `}
                                        onClick={(): void => this.setState({
                                            selectedButton: ButtonActions.Abstencion,
                                            errorMsg: false
                                        })}
                                    >
                                        <Radio
                                            value={ButtonActions.Abstencion}
                                            checked={selectedButton === ButtonActions.Abstencion} />
                                        <span className={styles.voteSingleButtonText}>
                                            {strings.Abstention}
                                        </span>
                                        <HandRightRegular />
                                    </Button>
                                    {/* En contra */}
                                    <Button
                                        className={`
                                        ${styles.voteSingleButton} 
                                        ${styles.voteSingleButtonRed}                                        
                                        ${(selectedButton === ButtonActions.EnContra) ?
                                                styles.selectedButton
                                                :
                                                styles.dimmedButton
                                            }
                                    `}
                                        onClick={(): void => this.setState({
                                            selectedButton: ButtonActions.EnContra,
                                            errorMsg: false
                                        })}
                                    >
                                        <Radio
                                            value={ButtonActions.EnContra}
                                            checked={selectedButton === ButtonActions.EnContra} />
                                        <span className={styles.voteSingleButtonText}>
                                            {strings.Reject}
                                        </span>
                                        <DismissCircleRegular />
                                    </Button>
                                    {
                                        /* Mensaje de error */
                                        (errorMsg) &&
                                        <div className={styles.voteErrorMessage}>
                                            <ErrorCircleRegular className={styles.voteErrorMessageIcon} />
                                            <span>
                                                {strings.MeetingVotingMustSelectOption}
                                            </span>
                                        </div>
                                    }
                                </RadioGroup>
                                {/* Botones de acción */}
                                <DialogActions className={styles.voteActionButtons}>
                                    <div className={styles.voteActionFirstRow}>
                                        {/* Botón para deseleccionar una opción */}
                                        <Button
                                            className={styles.voteActionSingleButton}
                                            appearance="secondary"
                                            disabled={selectedButton === ""}
                                            onClick={(): void => this.setState({
                                                selectedButton: "-1",
                                            })}
                                        >
                                            {strings.DeselectOption}
                                        </Button>
                                    </div>
                                    <div className={styles.voteActionSecondRow}>
                                        {/* Botón para cerrar */}
                                        <DialogTrigger disableButtonEnhancement>
                                            <Button
                                                className={styles.voteActionSingleButton}
                                                appearance="secondary"
                                                onClick={this.closeVoteModal}
                                            >
                                                {strings.RuleOut}
                                            </Button>
                                        </DialogTrigger>
                                        {/* Botón para guardar */}
                                        <Button
                                            className={styles.voteActionSingleButton}
                                            appearance="primary"
                                            onClick={this.saveFirstStepVoteModal}
                                        >
                                            {strings.ConfirmVote}
                                        </Button>
                                    </div>
                                </DialogActions>
                            </DialogBody>
                        }
                        {
                            /* Segunda pantalla - Voto seleccionado */
                            (secondScreen && canVote && !isLoading && !votingFinishedAndUserDidNotVote) &&
                            <div className={`${styles.voteBody} ${styles.voteBodySecondScreen}`}>
                                <div className={styles.voteBodySecondScreenContainer}>
                                    {/* Opción votada */}
                                    <div className={styles.voteOptions}>
                                        {/* Mensaje informativo */}
                                        <div className={styles.voteSelectedMessage}>
                                            {
                                                (this.testvVote === "") && (strings.ConfirmChosenOption + ":")
                                            }
                                            {
                                                (this.testvVote !== "") && (strings.VoteAlreadyRegisteredByTheUser + ":")
                                            }
                                        </div>
                                        <Button
                                            className={`
                                            ${styles.voteSingleButton}
                                            ${styles.selectedButton} 
                                            ${styles.noAction}
                                            ${(selectedButton === ButtonActions.AFavor) ?
                                                    styles.voteSingleButtonGreen
                                                    :
                                                    ''
                                                }
                                            ${(selectedButton === ButtonActions.Abstencion) ?
                                                    styles.voteSingleButtonGrey
                                                    :
                                                    ''
                                                }
                                            ${(selectedButton === ButtonActions.EnContra) ?
                                                    styles.voteSingleButtonRed
                                                    :
                                                    ''
                                                }
                                        `}
                                        >
                                            <Radio checked />
                                            {
                                                (selectedButton === ButtonActions.AFavor) &&
                                                <>
                                                    <span className={styles.voteSingleButtonText}>
                                                        {strings.Approdve}
                                                    </span>
                                                    <CheckmarkRegular />
                                                </>
                                            }
                                            {
                                                (selectedButton === ButtonActions.Abstencion) &&
                                                <>
                                                    <span className={styles.voteSingleButtonText}>
                                                        {strings.Abstention}
                                                    </span>
                                                    <HandRightRegular />
                                                </>
                                            }
                                            {
                                                (selectedButton === ButtonActions.EnContra) &&
                                                <>
                                                    <span className={styles.voteSingleButtonText}>
                                                        {strings.Reject}
                                                    </span>
                                                    <DismissCircleRegular />
                                                </>
                                            }
                                        </Button>
                                        {
                                            /* Mensaje para solicitar cambiar el voto */
                                            (this.testvVote !== "") &&
                                            <div className={styles.voteErrorMessage}>
                                                <ErrorCircleRegular className={styles.voteErrorMessageIcon} />
                                                <span>
                                                    {strings.ContactToChangeYourVote}
                                                </span>
                                            </div>
                                        }
                                    </div>
                                </div>
                                {/* Botones de acción */}
                                <DialogActions className={styles.voteActionButtons}>
                                    {
                                        (this.testvVote === "") &&
                                        <div className={styles.voteActionSecondRow}>
                                            {/* Botón para ir atrás/cerrar */}
                                            <DialogTrigger disableButtonEnhancement>
                                                <Button
                                                    className={styles.voteActionSingleButton}
                                                    appearance="secondary"
                                                    onClick={this.returnToFirstScreenVoteModal}
                                                >
                                                    {strings.Back}
                                                </Button>
                                            </DialogTrigger>
                                            {/* Botón para guardar */}
                                            <Button
                                                className={styles.voteActionSingleButton}
                                                appearance="primary"
                                                onClick={this.saveAndCloseVoteModal}
                                            >
                                                {strings.SubmitVote}
                                            </Button>
                                        </div>
                                    }
                                    {
                                        this.testvVote !== "" &&
                                        <div className={styles.voteActionSecondRow}>
                                            <DialogTrigger disableButtonEnhancement>
                                                <Button
                                                    className={styles.voteActionSingleButton}
                                                    appearance="primary"
                                                    onClick={this.closeVoteModal}>
                                                    {strings.Back}
                                                </Button>
                                            </DialogTrigger>
                                        </div>
                                    }

                                </DialogActions>
                            </div>
                        }
                        {
                            /* Pantalla para indicar que la votación ha finalizado y el usuario no votó */
                            (canVote && !isLoading && votingFinishedAndUserDidNotVote) &&
                            <div className={styles.voteVotingHasEndedAndYouDidNotVote}>
                                {/* Mensaje */}
                                <span className={styles.voteUserDidNotVoteMessage}>
                                    {strings.VotingHasEndedAndYouDidNotVote}
                                </span>
                                {/* Botón para cerrar */}
                                <div className={styles.voteUserDidNotVoteButtonRow}>
                                    <Button
                                        className={styles.voteUserDidNotVoteButton}
                                        appearance="primary"
                                        onClick={this.closeVoteModal}
                                    >
                                        {strings.RuleOut}
                                    </Button>
                                </div>
                            </div>
                        }
                        {
                            /* Pantalla para indicar que el usuario no puede votar */
                            (!canVote && !isLoading) &&
                            <div className={styles.voteNoPermission}>
                                {/* Mensaje */}
                                <span className={styles.voteNoPermissionMessage}>
                                    {strings.NoPermissionToVote}
                                </span>
                                {/* Botón para cerrar */}
                                <div className={styles.voteNoPermissionButtonRow}>
                                    <Button
                                        className={styles.voteNoPermissionButton}
                                        appearance="primary"
                                        onClick={this.closeVoteModal}
                                    >
                                        {strings.RuleOut}
                                    </Button>
                                </div>
                            </div>
                        }
                    </DialogSurface>
                </Dialog>
            );
        } else return (<div></div>)
    }

}