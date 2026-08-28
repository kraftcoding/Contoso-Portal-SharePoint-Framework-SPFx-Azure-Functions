import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'AgendaAnnouncementWebPartStrings';
import AgendaAnnouncement from './components/AgendaAnnouncement';
import SPService from '../../service/Service';
import { Version } from '@microsoft/sp-core-library';
import { IpropertyPaneConfiguration, propertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IAgendaAnnouncementprops } from './components/IAgendaAnnouncementprops';
import { Fluentprovider, Fluentproviderprops, Idprefixprovider, webLightTheme } from '@fluentui/react-components';
import { BackendService } from '../../service/BackendService';

export interface IAgendaAnnouncementWebPartprops {
  title: string;
}

export default class AgendaAnnouncementWebPart extends BaseClientSideWebPart<IAgendaAnnouncementWebPartprops> {

  public render(): void {
    const element: React.ReactElement<IAgendaAnnouncementprops> = React.createElement(
      AgendaAnnouncement,
      {
        title: this.properties.title,
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        bkService: this.context.serviceScope.consume(BackendService.serviceKey)        
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
        value: 'AgendaAnnouncement-'
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
                })
              ]
            }
          ]
        }
      ]
    };
  }

}