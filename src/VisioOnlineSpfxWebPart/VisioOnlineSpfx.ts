import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";

require('VisioEmbed');

import * as strings from 'VisioOnlineSpfxStrings';

import { VisioOnlineSpfxWebPart } from './VisioOnlineSpfxWebPart';

export interface IVisioOnlineScriptProps {
  url: string;
  width: string;
  height: string;
}

export default class VisioOnlineScript extends BaseClientSideWebPart<IVisioOnlineScriptProps> {

  public render(): void {

    const element = React.createElement(
      VisioOnlineSpfxWebPart,
      {
        url: this.properties.url,
        width: this.properties.width,
        height: this.properties.height
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [

                PropertyPaneTextField('width', {
                  label: 'Width',
                  value: '100%',
                  disabled: false,
                }),

                PropertyPaneTextField('height', {
                  label: 'Height',
                  value: '500px',
                  disabled: false,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
