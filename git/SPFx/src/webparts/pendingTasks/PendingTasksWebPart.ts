import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'PendingTasksWebPartStrings';
import SPService from '../../service/Service';
import { BackendService } from '../../service/BackendService';
import PendingTasks from './components/PendingTasks';
import { IPendingTasksprops } from './components/IPendingTasks';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { propertyPaneToggle, type IpropertyPaneConfiguration, propertyPaneTextField, propertyPaneSlider } from '@microsoft/sp-property-pane';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { providers, SharePointprovider } from '@microsoft/mgt-spfx';

export interface IPendingTasksWebPartprops {
  title: string;
  moreLink: string;
  numberOfResults: number;
  getThePendingTasksFromThisSite: boolean;
}

export default class PendingTasksWebPart extends BaseClientSideWebPart<IPendingTasksWebPartprops> {

  protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const pendingTasksprops: IPendingTasksprops = {
      context: this.context,
      spService: this.context.serviceScope.consume(SPService.serviceKey),
      bkService: this.context.serviceScope.consume(BackendService.serviceKey),
      locale: this.context.pageContext.cultureInfo.currentUICultureName,
      title: this.properties.title,
      moreLink: this.properties.moreLink,
      numberOfResults: this.properties.numberOfResults || 25,
      getThePendingTasksFromThisSite: this.properties.getThePendingTasksFromThisSite
    };
    const element: React.ReactElement<IPendingTasksprops> = React.createElement(
      PendingTasks,
      pendingTasksprops
    );
    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      { theme: webLightTheme },
      element
    );
    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      { value: "PendingTasks-" },
      fluentElement
    );
    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getpropertyPaneConfiguration(): IpropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                propertyPaneTextField("title", {
                  disabled: false,
                  label: strings.Title,
                  value: this.properties.title
                }),
                propertyPaneTextField("moreLink", {
                  label: strings.MoreLinkLabel,
                  placeholder: "https://"
                }),
                propertyPaneSlider("numberOfResults", {
                  label: strings.NumberOfResults,
                  min: 1,
                  max: 100,
                  value: this.properties.numberOfResults || 25
                }),
                propertyPaneToggle("getThePendingTasksFromThisSite", {
                  key: "getThePendingTasksFromThisSite",
                  label: strings.GetPendingTasksFromThisSite,
                  checked: this.properties.getThePendingTasksFromThisSite
                })
              ]
            }
          ]
        }
      ]
    };
  }

}