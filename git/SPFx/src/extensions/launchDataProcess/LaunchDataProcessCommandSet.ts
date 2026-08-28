import {Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
} from '@microsoft/sp-listview-extensibility';
import { IConfirmprocessDialogprops } from './ConfirmprocessDialog/IConfirmprocessDialog';
import ConfirmprocessDialog from './ConfirmprocessDialog/ConfirmprocessDialog';
import React from 'react';
import ReactDOM from 'react-dom';
import { BackendService } from '../../service/BackendService';
import SPService from '../../service/Service';

/**
 * If your command set uses the ClientSideComponentproperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ILaunchDataprocessCommandSetproperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'LaunchDataprocessCommandSet';

export default class LaunchDataprocessCommandSet extends BaseListViewCommandSet<ILaunchDataprocessCommandSetproperties> {
  private dialogPlaceHolder: HTMLDivElement;

  public onInit(): promise<void> {
    Log.info(LOG_SOURCE, 'Initialized LaunchDataprocessCommandSet');
    // initial state of the command's visibility
    const incrementalCommand: Command = this.tryGetCommand('INCREMENTAL');
    const completeCommand: Command = this.tryGetCommand('COMPLETE');
    incrementalCommand.visible = false;
    completeCommand.visible = false;

    const libraryUrl = this.context.pageContext.list?.serverRelativeUrl; 
    console.log(libraryUrl);
    if(libraryUrl && libraryUrl.toLowerCase() === '/sites/Contoso/processreports'){
      incrementalCommand.visible = true;
      completeCommand.visible = true;
      this.dialogPlaceHolder = document.body.appendChild(document.createElement("div"));
    }
    return promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'INCREMENTAL':
        this.renderConfirmDialog(false,true);
        break;
      case 'COMPLETE':
        this.renderConfirmDialog(true,true);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private renderConfirmDialog(isComplete:boolean, isOpen?:boolean){
    const element: React.ReactElement<IConfirmprocessDialogprops> = React.createElement(ConfirmprocessDialog, {
      isOpen: isOpen,
      onClosePanel: this.onCloseDialog,
      context: this.context,
      spService: this.context.serviceScope.consume(SPService.serviceKey),
      bkService: this.context.serviceScope.consume(BackendService.serviceKey),
      isComplete: isComplete,
    });
    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      {
        theme: webLightTheme,
      },
      element
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      {
        value: "Confirmprocess-",
      },
      fluentElement
    );
    ReactDOM.render(idprefixprovider, this.dialogPlaceHolder);
  }
  private onCloseDialog = (isComplete:boolean): void => {
    this.renderConfirmDialog(isComplete, false);
  };
  
/*
  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    var Libraryurl = this.context.pageContext.list.title; 
    const compareOneCommand: Command = this.tryGetCommand('INCREMENTAL');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
    */
}
