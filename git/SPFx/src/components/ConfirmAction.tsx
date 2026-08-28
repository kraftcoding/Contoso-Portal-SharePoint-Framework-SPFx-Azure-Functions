import * as React from 'react';
import {
    Button,
    Dialog,
    DialogSurface,
    DialogBody,
    DialogTitle,
    DialogContent,
    DialogActions,
    DialogTrigger
} from '@fluentui/react-components';
import { IConfirmActionprops } from './IConfirmActionprops';


export default class ConfirmAction extends React.Component<IConfirmActionprops, {}> {

 
    
    public render(): React.ReactElement<IConfirmActionprops> {
        const { acceptButtonprops, cancelButtonprops, dialogContent, dialogTitle, dialogTrigger, dialogprops, loadingActionContent } = this.props;
        //TODO: Find a better way. Fix to clip Sharepoint header bar in mobile.
        if(window.innerWidth < 640){
            let styles = document.createElement("style");
            styles.setAttribute("type", "text/css");
            styles.textContent = `#spSiteHeader{overflow-x: clip}`;
            document.head.appendChild(styles);
        }
     
        return (
            <Dialog {...dialogprops}>
                {
                    dialogTrigger &&
                    <DialogTrigger disableButtonEnhancement>
                        {dialogTrigger}
                    </DialogTrigger>
                }
                <DialogSurface>
                    <DialogBody>
                        <DialogTitle>{dialogTitle}</DialogTitle>
                        <DialogContent>
                            {dialogContent}
                        </DialogContent>
                        <DialogActions>
                            {loadingActionContent}
                            <DialogTrigger disableButtonEnhancement>
                                <Button {...cancelButtonprops}> {cancelButtonprops.title} </Button>
                            </DialogTrigger>
                            <DialogTrigger disableButtonEnhancement>
                                <Button {...acceptButtonprops}> {acceptButtonprops.title} </Button>
                            </DialogTrigger>
                        </DialogActions>
                    </DialogBody>
                </DialogSurface>
            </Dialog>
        );
    }

}