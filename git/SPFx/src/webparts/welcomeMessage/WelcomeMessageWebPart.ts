import * as React from 'react';
import * as ReactDom from 'react-dom';
import WelcomeMessage from './components/WelcomeMessage';
import SPService from '../../service/Service';
import { BackendService } from '../../service/BackendService';
import { IWelcomeMessageprops } from './components/IWelcomeMessage';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { propertyPaneToggle, type IpropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { Fluentprovider, Fluentproviderprops, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import strings from 'WelcomeMessageWebPartStrings';

export interface IWelcomeMessageWebPartprops {
  getNotificationsFromAllSites: boolean;
}

export default class WelcomeMessageWebPart extends BaseClientSideWebPart<IWelcomeMessageWebPartprops> {

  public render(): void {
    const element: React.ReactElement<IWelcomeMessageprops> = React.createElement(
      WelcomeMessage,
      {
        context: this.context,
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        bkService: this.context.serviceScope.consume(BackendService.serviceKey),
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
        getNotificationsFromAllSites: this.properties.getNotificationsFromAllSites
      }
    );

    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      {
        theme: webLightTheme
      },
      element
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      {
        value: 'WelcomeMessage-'
      },
      fluentElement
    )

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
                propertyPaneToggle("getNotificationsFromAllSites", {
                  key: "getNotificationsFromAllSites",
                  label: strings.GetNotificationsFromThisSite,
                  checked: this.properties.getNotificationsFromAllSites
                })
              ]
            }
          ]
        }
      ]
    };
  }

}