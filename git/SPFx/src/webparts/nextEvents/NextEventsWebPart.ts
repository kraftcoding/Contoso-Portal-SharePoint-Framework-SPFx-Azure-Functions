import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'NextEventsWebPartStrings';
import NextEvents from './components/NextEvents';
import SPService from '../../service/Service';
import { INextEventsprops } from './components/INextEventsprops';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IpropertyPaneConfiguration, propertyPaneLabel, propertyPaneTextField, propertyPaneSlider } from '@microsoft/sp-property-pane';
import { Fluentproviderprops, Fluentprovider, webLightTheme, Idprefixprovider } from '@fluentui/react-components';
import { BackendService } from '../../service/BackendService';

export interface INextEventsWebPartprops {
  title: string;
  slidesToShow: number;
  slidesToScroll: number;
  moreLink: string;
}

export default class NextEventsWebPart extends BaseClientSideWebPart<INextEventsWebPartprops> {

  private slidesToShow_minValue: number = 1;
  private slidesToShow_maxValue: number = 4;
  private slidesToScroll_minValue: number = 1;
  private slidesToScroll_maxValue: number = 4;

  public render(): void {
    const element: React.ReactElement<INextEventsprops> = React.createElement(
      NextEvents,
      {
        spService: this.context.serviceScope.consume(SPService.serviceKey),
        context: this.context,
        title: this.properties.title,
        slidesToShow: (this.properties.slidesToShow >= this.slidesToShow_minValue && this.properties.slidesToShow <= this.slidesToShow_maxValue) ? this.properties.slidesToShow : 0,
        slidesToScroll: (this.properties.slidesToScroll >= this.slidesToScroll_minValue && this.properties.slidesToScroll <= this.slidesToScroll_maxValue) ? (this.properties.slidesToScroll <= this.properties.slidesToShow ? this.properties.slidesToScroll : 0) : 0,
         moreLink: this.properties.moreLink,
        locale: this.context.pageContext.cultureInfo.currentUICultureName,
        bkService: this.context.serviceScope.consume(BackendService.serviceKey),
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
        value: 'NextEvents-'
      },
      fluentElement
    )
    ReactDom.render(idprefixprovider, this.domElement);
  }

  protected onInit(): promise<void> {
    return super.onInit();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getpropertyPaneConfiguration(): IpropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.Settingproperties,
              groupFields: [
                propertyPaneTextField("title", {
                  disabled: false,
                  label: strings.Title,
                  value: this.properties.title
                }),
                propertyPaneSlider("slidesToShow", {
                  disabled: false,
                  label: strings.SlidesToShow,
                  value: this.properties.slidesToShow,
                  min: this.slidesToShow_minValue,
                  max: this.slidesToShow_maxValue,
                  showValue: true
                }),
                propertyPaneSlider("slidesToScroll", {
                  disabled: false,
                  label: strings.SlidesToScroll,
                  value: this.properties.slidesToScroll,
                  min: this.slidesToScroll_minValue,
                  max: this.slidesToScroll_maxValue,
                  showValue: true
                }),
                propertyPaneTextField('moreLink', {
                  label: strings.MoreLinkLabel,
                  placeholder: 'https://'
                }),
                propertyPaneLabel("", {
                  text: strings.SlideInformation
                })
              ]
            }
          ]
        }
      ]
    };
  }

}