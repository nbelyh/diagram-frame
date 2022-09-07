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
import { PropertyPaneSizeField } from './PropertyPaneSizeField';

export interface IWebPartProps {
  url: string;
  startPage: string;
  width: string;
  height: string;
  hideToolbars: boolean;
  hideBorders: boolean;

  hideDiagramBoundary: boolean;
  disableHyperlinks: boolean;
  disablePan: boolean;
  disablePanZoomWindow: boolean;
  disableZoom: boolean;

  zoom: number;
}

export default class WebPart extends BaseClientSideWebPart<IWebPartProps> {

  public onInit(): Promise<void> {

    return super.onInit().then(() => {
      sp.setup({ spfxContext: this.context as any });
    });
  }

  public render(): void {

    const isPropertyPaneOpen = this.context.propertyPane.isPropertyPaneOpen();

    const element: React.ReactElement = (this.properties.url)
      ? React.createElement(TopFrame, {
        ...this.properties,
        context: this.context
      })
      : React.createElement(Placeholder, {
        iconName: "Edit",
        iconText: isPropertyPaneOpen ? "Select Visio Diagram" : "Configure Web Part",
        description: isPropertyPaneOpen ? "Click 'Browse...' Button on configuration panel to select the diagram" : "Press 'Configure' button to configure the web part",
        buttonLabel: "Configure",
        onConfigure: () => this.context.propertyPane.open(),
        hideButton: isPropertyPaneOpen
      });

    ReactDom.render(element, this.domElement);
  }

  public onPropertyPaneConfigurationStart() {
    this.render();
  }

  public onPropertyPaneConfigurationComplete() {
    this.render();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: "Drawing Display",
              groupFields: [
                PropertyPaneUrlField('url', {
                  url: this.properties.url,
                  context: this.context
                }),

                PropertyPaneTextField('startPage', {
                  label: strings.FieldStartPage,
                  description: "Page (name) to activate on load"
                }),

                PropertyPaneTextField('zoom', {
                  label: strings.FieldZoom,
                  description: "Zoom level (percents) to set on load"
                }),
              ]
            },
            {
              groupName: "Appearance",
              groupFields: [
                PropertyPaneSizeField('width', {
                  label: strings.FieldWidth,
                  description: "Specify value and units (leave blank for default)",
                  value: this.properties.width,
                  screenUnits: 'w'
                }),

                PropertyPaneSizeField('height', {
                  label: strings.FieldHeight,
                  description: "Specify value and units (leave blank for default)",
                  value: this.properties.height,
                  screenUnits: 'h'
                }),
                PropertyPaneToggle('hideToolbars', {
                  label: "Hide Toolbars",
                }),
                PropertyPaneToggle('hideDiagramBoundary', {
                  label: "Hide Diagram Boundary",
                }),
                PropertyPaneToggle('hideBorders', {
                  label: "Hide Borders",
                }),
              ]
            },
            {
              groupName: "Drawing Interactivity",
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('disableHyperlinks', {
                  label: "Disable Hyperlinks",
                }),
                PropertyPaneToggle('disablePan', {
                  label: "Disable Pan",
                }),
                PropertyPaneToggle('disableZoom', {
                  label: "Disable Zoom",
                }),
                PropertyPaneToggle('disablePanZoomWindow', {
                  label: "Disable PanZoom Window",
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
