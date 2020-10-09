import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { sp } from '@pnp/sp';

require('VisioEmbed');

import * as strings from 'WebPartStrings';

import { TopFrame } from './TopFrame';
import { PropertyPaneVersionField } from './PropertyPaneVersionField';
import { PropertyPaneUrlField } from './PropertyPaneUrlField';
import { Placeholder } from '@pnp/spfx-controls-react';

export interface IVisioOnlineScriptProps {
  width: string;
  height: string;
  showToolbars: boolean;
  showBorders: boolean;
  url: string;
  zoom: number;
}

export default class WebPart extends BaseClientSideWebPart<IVisioOnlineScriptProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({ spfxContext: this.context });
    });
  }

  public render(): void {

    const element: React.ReactElement = (this.properties.url)
      ? React.createElement(TopFrame, {
        width: this.properties.width,
        height: this.properties.height,
        context: this.context,
        url: this.properties.url,
        showToolbars: this.properties.showToolbars,
        showBorders: this.properties.showBorders,
        zoom: +this.properties.zoom
      })
      : React.createElement(Placeholder, {
        iconName: "Edit",
        iconText: "Select Diagram",
        description: "Press 'Configure' button to choose Visio file to display here",
        buttonLabel: "Configure",
        onConfigure: () => this.context.propertyPane.open()
      });

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
                PropertyPaneUrlField('url', {
                  url: this.properties.url,
                  context: this.context
                }),

                PropertyPaneTextField('width', {
                  label: strings.FieldWidth,
                }),

                PropertyPaneTextField('height', {
                  label: strings.FieldHeight,
                }),

                PropertyPaneTextField('zoom', {
                  label: strings.FieldZoom,
                }),
              ]
            },
            {
              groupName: strings.Toolbars,
              groupFields: [
                PropertyPaneToggle('showToolbars', {
                  label: strings.FieldShowToolbars,
                }),

                PropertyPaneToggle('showBorders', {
                  label: strings.FieldShowBorders,
                }),
              ]
            },
            {
              groupName: "About",
              groupFields: [
                PropertyPaneVersionField(this.context.manifest.version)
              ]
            }
          ]
        }
      ]
    };
  }
}
