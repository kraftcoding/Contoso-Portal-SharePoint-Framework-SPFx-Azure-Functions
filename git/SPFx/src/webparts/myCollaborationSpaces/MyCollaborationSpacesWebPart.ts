import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IpropertyPaneConfiguration, propertyPaneToggle, propertyPaneTextField, propertyPaneSlider } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import MyCollaborationSpaces from './components/MyCollaborationSpaces';
import { IMyCollaborationSpacesprops } from './components/IMyCollaborationSpacesprops';
import SPService from '../../service/Service';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
import * as strings from 'MyCollaborationSpacesWebPartStrings';

export interface IMyCollaborationSpacesWebPartprops {
  title: string;
  numberOfResults: number;
  showDescription: boolean;
  moreLink: string;
}

export default class MyCollaborationSpacesWebPart extends BaseClientSideWebPart<IMyCollaborationSpacesWebPartprops> {

  public render(): void {
    const element: React.ReactElement<IMyCollaborationSpacesprops> = React.createElement(
      MyCollaborationSpaces,
      {
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        title: this.properties.title,
        numberOfResults: this.properties.numberOfResults,
        showDescription: this.properties.showDescription,
        moreLink: this.properties.moreLink,
        context: this.context
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
        value: 'MyCollabSpaces-'
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
                propertyPaneTextField('title', {
                  label: strings.Title
                }),
                propertyPaneTextField('moreLink', {
                  label: strings.MoreLinkLabel,
                  placeholder: 'https://'
                }),
                propertyPaneSlider('numberOfResults', {
                  disabled: false,
                  label: strings.NumberOfResults,
                  min: 1,
                  max: 12,
                  value: this.properties.numberOfResults
                }),
                propertyPaneToggle('showDescription', {
                  checked: this.properties.showDescription,
                  label: strings.ShowDescription,
                  onText: strings.Yes,
                  offText: strings.No
                })
              ]
            }
          ]
        }
      ]
    };
  }

}