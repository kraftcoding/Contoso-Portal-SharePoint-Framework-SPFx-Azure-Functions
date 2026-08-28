import * as React from 'react';
import * as strings from 'MeetingsWebPartStrings';
import styles from './ShowQRCode.module.scss';
import {
    IShowQRCodeprops,
    IShowQRCodeState
} from '../IMeetingVotingModals';
import {
    VotingContext,
    VotingDialogActions
} from '../../IMeetingVoting';
import {
    DismissFilled,
    QrCodeRegular
} from '@fluentui/react-icons';
import {
    Button,
    Dialog,
    DialogActions,
    DialogBody,
    DialogContent,
    DialogSurface,
    DialogTrigger,
    Spinner
} from '@fluentui/react-components';
import { Logger } from '../../../../../../utils/Logger';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import QRCode from 'qrcode'

export default class ShowQRCode extends React.Component<IShowQRCodeprops, IShowQRCodeState> {

    static contextType = VotingContext;
    context!: React.ContextType<typeof VotingContext>;

    constructor(props: IShowQRCodeprops) {
        super(props);
        this.state = {
            qrUrl: "",
            qrcodedata: "",
            errorMsg: ""
        }
    }

    public async componentDidMount(): promise<void> {
        await this.generateQr();
    }

    public async componentDidUpdate(testvprops: Readonly<IShowQRCodeprops>): promise<void> {
        if (testvprops.votingSelected !== this.props.votingSelected) {
            await this.generateQr();
        }
    }

    private async generateQr(): promise<void> {
        const {
            votingSelected
        } = this.props;
        const {
            event,
            context
        } = this.context;

        const qrUrl: string = context.pageContext.web.absoluteUrl + "#/" + event.Id + "/Vote/" + votingSelected?.Id;
        const errorContext: WebPartContext = this.context.context;

        try {
            if (qrUrl) {
                const qrcodedata: string = await QRCode.toDataURL(qrUrl);
                this.setState({
                    qrUrl,
                    qrcodedata
                });
            }
        }
        catch (error) {
            Logger.error("Error - ShowQRCode - componentDidMount - Error generating QR code", error, errorContext);
            this.setState({
                errorMsg: `${strings.ErrorToCreateQRCode} ${error}`
            });
        }
    }

    public render(): React.ReactElement<IShowQRCodeprops> {
        const {
            qrcodedata
        } = this.state;
        const {
            event,
            context,
            votingDialog,
            votingSelected,
            onDismissModal
        } = this.context;

        const dialogIsOpen: boolean = (votingDialog === VotingDialogActions.ShowQRCode);
        const DepartmentName: string = (context?.pageContext.web.title) ?? "";

        return (
            <Dialog open={dialogIsOpen}>
                <DialogSurface className={styles.showQRCode}>
                    {/* Cabecera */}
                    <div className={styles.showQRCodeHeader}>
                        {/* Icono */}
                        <div className={styles.showQRCodeLogo}>
                            <QrCodeRegular />
                        </div>
                        {/* Título del órgano */}
                        <div className={styles.showQRCodeTitle}>
                            {DepartmentName}
                        </div>
                        {/* Botón de cerrar */}
                        <Button
                            icon={<DismissFilled />}
                            appearance="transparent"
                            className={styles.showQRCodeDismissIcon}
                            onClick={onDismissModal}
                        />
                    </div>
                    {/* Subcabecera */}
                    <div className={styles.showQRCodeSubheader}>
                        {/* Título de la Meeting */}
                        <div className={styles.showQRCodeFirstText}>
                            {event.Title}
                        </div>
                        {/* Título de la votación */}
                        <div className={styles.showQRCodeSecondText}>
                            {
                                (votingSelected?.Order && votingSelected?.Title) ?
                                    (votingSelected.Order + "- " + votingSelected.Title)
                                    :
                                    undefined
                            }
                        </div>
                    </div>
                    {/* Contenido */}
                    <DialogBody className={styles.showQRCodeBody}>
                        <DialogContent className={styles.showQRCodeBodyContent}>
                            {
                                (qrcodedata) ?
                                    /* Imagen del código QR */
                                    <div className={styles.qrCodeImageContainer}>
                                        <img className={styles.qrCodeImage} src={qrcodedata} alt={strings.QRCode} />
                                    </div>
                                    :
                                    /* Generando el código QR */
                                    <div className={styles.generatingTheQRCode}>
                                        <Spinner /> <span> {strings.GeneratingTheQRCode + "..."} </span>
                                    </div>
                            }
                        </DialogContent>
                        <DialogActions className={styles.showQRCodeActionButtons}>
                            {/* Botón para cerrar */}
                            <DialogTrigger disableButtonEnhancement>
                                <Button
                                    className={styles.showQRCodeActionSingleButton}
                                    appearance="secondary"
                                    onClick={onDismissModal}
                                >
                                    {strings.Close}
                                </Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

}