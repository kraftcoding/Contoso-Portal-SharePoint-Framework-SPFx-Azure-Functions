import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IpropertyPaneConfiguration,
  propertyPaneTextField,
  propertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MasterTableStatusWebPartStrings';
import MasterTableStatus from './components/MasterTableStatus';
import { IMasterTableStatusprops } from './components/IMasterTableStatusprops';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { providers, SharePointprovider } from '@microsoft/mgt-spfx';


export interface IMasterTableStatusWebPartprops {
  title: string;
  pageSize: number;
}

export default class MasterTableStatusWebPart extends BaseClientSideWebPart<IMasterTableStatusWebPartprops> {

  public render(): void {
    const element: React.ReactElement<IMasterTableStatusprops> = React.createElement(
      MasterTableStatus,
      {
        title: this.properties.title,
        context: this.context,
        pageSize: this.properties.pageSize ?? 5
      }
    );


    const fluentElement: React.ReactElement<Fluentproviderprops> = React.createElement(
      Fluentprovider,
      { theme: webLightTheme },
      element
    );

    const idprefixprovider: React.ReactElement<Fluentproviderprops> = React.createElement(
      Idprefixprovider,
      { value: 'MasterTableStatus-' },
      fluentElement
    )

    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected onInit(): promise<void> {
    if (!providers.globalprovider) {
      providers.globalprovider = new SharePointprovider(this.context);
    }
    return promise.resolve();
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
                propertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                })
              ]
            },
            {
              groupName: strings.PaginationGroupName,
              groupFields: [
                propertyPaneSlider('pageSize', {
                  label: strings.PageSizeFieldLabel,
                  min: 1,
                  max: 50,
                  step: 1,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}