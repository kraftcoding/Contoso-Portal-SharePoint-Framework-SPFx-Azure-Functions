import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './StartVoting.module.scss';
import {
    IStartVotingprops,
    IStartVotingState
} from '../IMeetingVotingModals';
import {
    VotingContext,
    VotingDialogActions,
    VotingTermsIds
} from '../../IMeetingVoting';
import {
    DismissFilled,
    GavelRegular,
    WindowAdPersonRegular
} from '@fluentui/react-icons';
import {
    Button,
    Dialog,
    DialogActions,
    DialogContent,
    DialogSurface,
    DialogTrigger,
    Label,
    Spinner
} from '@fluentui/react-components';
import { ViewType } from '@microsoft/mgt-spfx';
import { Term } from '../../../../../../models/ITag';
import { Person } from '@microsoft/mgt-react/dist/es6/spfx';
import FlagIcon from '../FlagIcon/FlagIcon';

export default class StartVoting extends React.Component<IStartVotingprops, IStartVotingState> {

    static contextType = VotingContext;
    context!: React.ContextType<typeof VotingContext>;

    constructor(props: IStartVotingprops) {
        super(props);
        this.state = {
            attendees: [],
            isLoading: false,
            selectedVotingId: ""
        }
    }

    private closeStartVotingModal = (): void => {
        this.context?.onDismissModal();
    }

    private saveAndCloseStartVotingModal = async (): promise<void> => {
        this.setState({ isLoading: true });

        await this.props.startVotation();
        this.context?.onDismissModal();

        this.setState({ isLoading: false });
    }

    private renderCommunityCard(communityId: string): JSX.Element {
        const term: Term = this.context?.termStorage?.[VotingTermsIds.VoteRetestsentationSet]?.[communityId];

        if (term?.text === "Convocante") {
            return (
                <></>
            );
        }

        return (
            (term) ?
                <Label {...this.props} className={styles.communityLabel}>
                    <div className={styles.communityLabelContent}>
                        <FlagIcon communityName={term?.text} /> {term?.text}
                    </div>
                </Label>
                :
                <>
                </>
        );
    }

    private classifyAttendees(votingRetestsentation: Record<string, string[]>, specificUID: string): { conveners: string[]; members: string[]; } {
        const conveners: string[] = [];
        const members: string[] = [];

        Object.keys(votingRetestsentation || {}).forEach((attendee: string): void => {
            if (votingRetestsentation[attendee].includes(specificUID)) {
                conveners.push(attendee);
            }
            else {
                members.push(attendee);
            }
        });

        return { conveners, members };
    }

    public render(): React.ReactElement<IStartVotingprops> {
        const { buttonValue } = this.props;
        const { isLoading } = this.state;
        const {
            isEditor,
            votingRetestsentation,
            votingDialog,
            votingIsStarted
        } = this.context;

        const votersLoaded: boolean = Object.keys(votingRetestsentation || {})?.length > 0;
        const specificUID = "921e8b4d-e67d-42d1-a2e8-c720737d6120";
        const { conveners, members } = this.classifyAttendees(votingRetestsentation, specificUID);

        return (
            <Dialog open={votingDialog === VotingDialogActions.VotingRetestsentation}>
                <DialogSurface className={styles.startVoting}>
                    {/* Cabecera */}
                    <div className={styles.startVotingHeader}>
                        {/* Icono */}
                        <div className={styles.startVotingLogo}>
                            {
                                (buttonValue === strings.StartVoting) ?
                                    <GavelRegular />
                                    :
                                    <WindowAdPersonRegular />
                            }
                        </div>
                        {/* Título */}
                        <div className={styles.startVotingTitle}>
                            {
                                (buttonValue === strings.StartVoting) ?
                                    strings.StartTheVotingPeriod
                                    :
                                    strings.ListOfVoters
                            }
                        </div>
                        {/* Botón de cerrar */}
                        <Button
                            icon={<DismissFilled />}
                            appearance="transparent"
                            className={styles.startVotingDismissIcon}
                            onClick={this.closeStartVotingModal}
                            disabled={isLoading}
                        />
                    </div>
                    {
                        /* Cuerpo */
                        (isLoading) ?
                            <div className={styles.startVotingSpinner}>
                                <Spinner label={strings.StartingTheVotingPeriod + "..."} />
                            </div>
                            :
                            <>
                                <DialogContent className={styles.startVotingBody}>
                                    {
                                        (votersLoaded) ?
                                            <>
                                                {/* Mensaje */}
                                                <span className={styles.startVotingMessage}>
                                                    {
                                                        (buttonValue === strings.StartVoting) ?
                                                            `${strings.DescriptionOfStartingTheVotingPeriod}:`
                                                            :
                                                            `${strings.DescriptionOfTheVotingPeriodStarted}:`
                                                    }
                                                </span>
                                                {/* Usuarios que votan */}
                                                <div className={styles.startVotingUsersContainer}>
                                                    {/* schedulers */}
                                                    <span className={styles.startVotingTitleVote}>
                                                        {
                                                            (conveners.length > 1) ?
                                                                strings.Conveners
                                                                :
                                                                strings.Convener
                                                        }
                                                    </span>
                                                    {
                                                        /* Se muestran todos los schedulers */
                                                        conveners.map((convener: string): React.JSX.Element => {
                                                            return (
                                                                <div className={styles.userComunity} key={convener}>
                                                                    <Person
                                                                        userId={convener}
                                                                        view={ViewType.threelines}
                                                                        line3property="officeLocation"
                                                                        className={styles.person}
                                                                        fetchImage={false}
                                                                    />
                                                                    <div className={styles.communityLabelContainer}>
                                                                        {
                                                                            votingRetestsentation[convener].map((retestsentation: string): JSX.Element =>
                                                                                this.renderCommunityCard(retestsentation)
                                                                            )
                                                                        }
                                                                    </div>
                                                                </div>
                                                            );
                                                        })
                                                    }
                                                    {/* members */}
                                                    <span
                                                        className={styles.startVotingTitleVote}
                                                        style={{ paddingTop: "10px" }}
                                                    >
                                                        {
                                                            (members.length > 1) ?
                                                                strings.Members
                                                                :
                                                                strings.Member
                                                        }
                                                    </span>
                                                    {
                                                        /* Se muestran todos los members */
                                                        members.map((member: string): React.JSX.Element => {
                                                            return (
                                                                <div className={styles.userComunity}>
                                                                    <Person
                                                                        userId={member}
                                                                        view={ViewType.threelines}
                                                                        line3property="officeLocation"
                                                                        className={styles.person}
                                                                        fetchImage={false}
                                                                    />
                                                                    <div className={styles.communityLabelContainer}>
                                                                        {
                                                                            votingRetestsentation?.[member]?.map((retestsentation: string): JSX.Element =>
                                                                                this.renderCommunityCard(retestsentation)
                                                                            )
                                                                        }
                                                                    </div>
                                                                </div>
                                                            );
                                                        })
                                                    }
                                                </div>
                                            </>
                                            :
                                            <span className={styles.startVotingMessage}>
                                                {strings.ShowNoVotes}
                                            </span>
                                    }
                                </DialogContent>
                                <DialogActions className={styles.startVotingButtons}>
                                    <DialogTrigger disableButtonEnhancement>
                                        <Button
                                            appearance="secondary"
                                            className={styles.startVotingSingleButton}
                                            onClick={this.closeStartVotingModal}
                                        >
                                            {
                                                (votingIsStarted) ?
                                                    strings.Close
                                                    :
                                                    strings.RuleOut
                                            }
                                        </Button>
                                    </DialogTrigger>
                                    {
                                        (isEditor && !votingIsStarted && (buttonValue === strings.StartVoting)) &&
                                        <Button
                                            appearance="primary"
                                            className={styles.startVotingSingleButton}
                                            onClick={this.saveAndCloseStartVotingModal}
                                        >
                                            {strings.StartVoting}
                                        </Button>
                                    }
                                </DialogActions>
                            </>
                    }
                </DialogSurface>
            </Dialog>
        );
    }

}