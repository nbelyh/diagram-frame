import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from "@microsoft/sp-property-pane";
import { sp } from '@pnp/sp';

require('VisioEmbed');

import * as strings from 'WebPartStrings';

import { TopFrame } from './TopFrame';
import { PropertyPaneVersionField } from './PropertyPaneVersionField';
import { PropertyPaneUrlField } from './PropertyPaneUrlField';
import { PropertyPaneSizeField } from './PropertyPaneSizeField';
import { IDefaultFolder } from './IDefaultFolder';

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

  private defaultFolder: IDefaultFolder;
  private async getDefaultFolder(): Promise<IDefaultFolder> {
    if (this.defaultFolder) {
      return this.defaultFolder;
    }

    const teamsContext = this.context.sdks.microsoftTeams?.context;
    if (teamsContext) {
      return this.defaultFolder = {
        name: teamsContext.channelName,
        relativeUrl: teamsContext.channelRelativeUrl
      };
    }

    try {
      const lists = await sp.web.lists.select('DefaultViewURL', 'Title').filter('BaseTemplate eq 101 and Hidden eq false').get();
      const firstList = lists[0];
      if (firstList) {
        const webUrl = this.context.pageContext.web.serverRelativeUrl;
        let viewUrl = firstList.DefaultViewUrl;
        if (viewUrl.startsWith(webUrl))
          viewUrl = viewUrl.substring(webUrl.length);

        const pos = viewUrl.indexOf("/Forms/");
        if (pos >= 0) {
          const docLibPath = viewUrl.substring(0, pos);
          return this.defaultFolder = {
            name: firstList.Title,
            relativeUrl: `${webUrl}${docLibPath}`
          }
        }
      }
    } catch (err) {
      console.warn('Unable to dtermine default folder using default', err);
    }

    return this.defaultFolder = {
      name: undefined,
      relativeUrl: undefined
    };
  }

  public onInit(): Promise<void> {

    return super.onInit().then(() => {
      sp.setup({ spfxContext: this.context as any });
    });
  }

  public render(): void {

    const isPropertyPaneOpen = this.context.propertyPane.isPropertyPaneOpen();

    const properties = {
      ...this.properties,
      width: this.properties.width || '100%',
      height: this.properties.height || '50vh'
    };

    const element = React.createElement(TopFrame, {
      ...properties,
      isPropertyPaneOpen,
      onConfigure: () => this.context.propertyPane.open(),
      isReadOnly: this.displayMode === DisplayMode.Read,
      context: this.context
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

  private defaultWidth;
  private async getDefaultWidth() {
    if (this.defaultWidth) {
      return this.defaultWidth;
    }

    return this.defaultWidth = '100%';
  }

  private defaultHeight;
  private async getDefaultHeight() {

    if (this.defaultHeight) {
      return this.defaultHeight;
    }

    if (this.context.sdks.microsoftTeams) {
      return this.defaultHeight = '100%';
    }

    const pageContext = this.context.pageContext;
    if (pageContext?.list?.id && pageContext?.listItem?.id) {
      try {
        const item = await sp.web.lists.getById(pageContext.list.id.toString()).items.getById(pageContext.listItem.id).select('PageLayoutType').get();
        if (item['PageLayoutType'] === 'SingleWebPartAppPage') {
          return this.defaultHeight = '100%';
        }
      } catch (err) {
        console.warn('Unable to dtermine default height using default', err);
      }
    }
    return this.defaultHeight = '50vh';
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
                  context: this.context,
                  getDefaultFolder: () => this.getDefaultFolder(),
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
                  description: "Specify width value and units",
                  value: this.properties.width,
                  screenUnits: 'w',
                  getDefaultValue: () => this.getDefaultWidth()
                }),

                PropertyPaneSizeField('height', {
                  label: strings.FieldHeight,
                  description: "Specify height value and units",
                  value: this.properties.height,
                  screenUnits: 'h',
                  getDefaultValue: () => this.getDefaultHeight()
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
