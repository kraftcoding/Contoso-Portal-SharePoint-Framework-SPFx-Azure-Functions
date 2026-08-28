import * as React from 'react';
import styles from './NewDocForm.module.scss';
import { INewDocFormprops } from './INewDocFormprops';
import { INewDocFormState } from './INewDocFormState';
import {
    Button,
    Dialog,
    DialogSurface,
    DialogBody,
    DialogTitle,
    DialogContent,
    DialogActions,
    Dropdown,
    Label,
    Option,
    Field,
    Spinner,
    Input
} from '@fluentui/react-components';
import { Logger } from '../utils/Logger';
import { IFileInfo } from '@pnp/sp/files';
import { find } from '@microsoft/sp-lodash-subset/lib/index';
import strings from 'MeetingsWebPartStrings';
import { sortByprodp } from '../utils/Utils';

export default class NewDocForm extends React.Component<INewDocFormprops, INewDocFormState> {

    constructor(props: INewDocFormprops) {
        super(props);
        this.state = {
            isLoading: false,
            isSaving: false,
            files: [],
            fileError: ''
        };
    }

    public componentDidMount(): void {
        void this.onInit();
    }

    private async onInit(): promise<void> {
        const { spService, relativeUrlToGetDocs } = this.props;

        this.setState({ isLoading: true });

        
        const filesRes: IFileInfo[] = await spService.getFiles(relativeUrlToGetDocs);
        // Sort files by property Name
        const files: IFileInfo[] = sortByprodp(filesRes, 'Name');

        this.setState({ isLoading: false, isSaving: false, files});
    }

    private onSave = async (): promise<void> => {
        const { onSaveData } = this.props;
        const { selectedFile } = this.state;

        if (selectedFile) {
            this.setState({ isSaving: true });

            try {
                await onSaveData(selectedFile.optionValue);
            } catch (e) {
                Logger.error(e);
            } finally {
                this.onDismiss();
            }
        } else {
            this.setState({ fileError: strings.SelectingADocumentIsRequired });
        }
    };

    private onDismiss = (): void => this.setState({ selectedFile: undefined, fileError: '', isLoading: false, isSaving: false }, this.props.openCloseForm);

    public render(): React.ReactElement<INewDocFormprops> {
        const { open, relatedItemTitle, relatedDocumentsIds } = this.props;
        const { isLoading, files, selectedFile, fileError, isSaving } = this.state;

        return (
            <Dialog open={open} onOpenChange={this.onDismiss}>
                <DialogSurface className={styles.dialogDocsSurface}>
                    <DialogBody>
                        <DialogTitle> {strings.AssociateDocument} </DialogTitle>
                        <DialogContent className={styles.dialogContent}>
                            <Label> {strings.Document} </Label>
                            <Field validationMessage={fileError}>
                                {files && <Dropdown
                                    disabled={isLoading || isSaving}
                                    className={styles.section}
                                    placeholder={strings.SelectADocument}
                                    value={selectedFile?.optionText || ''}
                                    selectedOptions={selectedFile ? [selectedFile.optionValue] : []}
                                    onOptionSelect={(ev, data) => this.setState({ selectedFile: data, fileError: '' })}
                                >
                                    {files.map((file: IFileInfo, i: number): JSX.Element => {
                                        return <Option disabled={find(relatedDocumentsIds, docId => docId === file.UniqueId) !== undefined} key={file.UniqueId} value={file.UniqueId}>{file.Name}</Option>;
                                    })}
                                </Dropdown>
                                }
                            </Field>
                            <Label> {strings.AssociateTo} </Label>
                            <Input
                                disabled
                                className={styles.section}
                                value={`${relatedItemTitle}`}
                            />
                        </DialogContent>
                        <DialogActions>
                            {isSaving && <Spinner size='tiny' />}
                            <Button disabled={isSaving} appearance='secondary' onClick={this.onDismiss}> {strings.RuleOut} </Button>
                            <Button disabled={isSaving} appearance='primary' onClick={this.onSave}> {strings.Save} </Button>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

}