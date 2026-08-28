import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IpropertyPaneConfiguration,
  propertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DelegateAdministrationBannerWebPartStrings';
import DelegateAdministrationBanner from './components/DelegateAdministrationBanner';
import { IDelegateAdministrationBannerprops } from './components/IDelegateAdministrationBanner';
import { providers, SharePointprovider } from '@microsoft/mgt-spfx';
import { BackendService } from '../../service/BackendService';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
import SPService from '../../service/Service';

export interface IDelegateAdministrationBannerWebPartprops {
  componentTitle: string;
}

export default class DelegateAdministrationBannerWebPart extends BaseClientSideWebPart<IDelegateAdministrationBannerWebPartprops> {
 protected async onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
  }

  public render(): void {
    const { context, properties } = this;
    
    const delegateprops: IDelegateAdministrationBannerprops = {
      context,
      spService: this.context.serviceScope.consume(SPService.serviceKey),
      bkService: context.serviceScope.consume(BackendService.serviceKey),
      componentTitle: properties.componentTitle,
    }
    const reactElement: React.ReactElement<IDelegateAdministrationBannerprops> = React.createElement(
          DelegateAdministrationBanner,
          delegateprops
        );
    
        const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
          Fluentprovider,
          { theme: webLightTheme },
          reactElement
        );
    
        const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
          Idprefixprovider,
          { value: 'AdministrationApp-' },
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
          header: {
            description: strings.propertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                propertyPaneTextField('componentTitle', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
