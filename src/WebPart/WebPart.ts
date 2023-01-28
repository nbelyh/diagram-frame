import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { sp } from '@pnp/sp';

require('VisioEmbed');

import * as strings from 'WebPartStrings';

import { TopFrame } from './TopFrame';
import { PropertyPaneVersionField } from './properties/PropertyPaneVersionField';
import { PropertyPaneUrlField } from './properties/PropertyPaneUrlField';
import { PropertyPaneSizeField } from './properties/PropertyPaneSizeField';
import { IDefaultFolder } from './properties/IDefaultFolder';
import { IWebPartProps } from './IWebPartProps';

export default class WebPart extends BaseClientSideWebPart<IWebPartProps> {

  private defaultFolder: IDefaultFolder;
  private async getDefaultFolder(): Promise<IDefaultFolder> {
    if (this.defaultFolder) {
      return this.defaultFolder;
    }

    const teamsContext = this.context.sdks?.microsoftTeams?.context;
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

        const pos = viewUrl.indexOf('/Forms/');
        if (pos >= 0) {
          const docLibPath = viewUrl.substring(0, pos);
          return this.defaultFolder = {
            name: firstList.Title,
            relativeUrl: `${webUrl}${docLibPath}`
          }
        }
      }
    } catch (err) {
      console.warn('Unable to determine default folder using default', err);
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
      context: this.context,
      isTeams: !!this.context.sdks?.microsoftTeams?.context
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
        console.warn('Unable to determine default height using default', err);
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
              groupName: strings.PropertyPaneLabelDrawingDisplay, // Drawing Display
              groupFields: [
                PropertyPaneUrlField('url', {
                  url: this.properties.url,
                  context: this.context,
                  getDefaultFolder: () => this.getDefaultFolder(),
                }),

                PropertyPaneTextField('startPage', {
                  label: strings.FieldStartPage,
                  description: strings.FieldStartPageDescription
                }),

                PropertyPaneTextField('zoom', {
                  label: strings.FieldZoom,
                  description: strings.FieldZoomDescription
                }),
                PropertyPaneToggle('enableNavigation', {
                  label: strings.FieldEnableNavigation,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelAppearance, // Appearance
              groupFields: [
                PropertyPaneSizeField('width', {
                  label: strings.FieldWidth,
                  description: strings.FieldWidthDescription,
                  value: this.properties.width,
                  screenUnits: 'w',
                  getDefaultValue: () => this.getDefaultWidth()
                }),

                PropertyPaneSizeField('height', {
                  label: strings.FieldHeight,
                  description: strings.FieldHeightDescription,
                  value: this.properties.height,
                  screenUnits: 'h',
                  getDefaultValue: () => this.getDefaultHeight()
                }),
                PropertyPaneToggle('hideToolbars', {
                  label: strings.PropertyPaneLabelhideToolbars,
                }),
                PropertyPaneToggle('hideDiagramBoundary', {
                  label: strings.PropertyPaneLabelhideDiagramBoundary,
                }),
                PropertyPaneToggle('hideBorders', {
                  label: strings.PropertyPaneLabelhideBorders,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelInteractivity, // Drawing Interactivity
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('disableHyperlinks', {
                  label: strings.PropertyPaneLabeldisableHyperlinks,
                }),
                PropertyPaneToggle('disablePan', {
                  label: strings.PropertyPaneLabeldisablePan,
                }),
                PropertyPaneToggle('disableZoom', {
                  label: strings.PropertyPaneLabeldisableZoom,
                }),
                PropertyPaneToggle('disablePanZoomWindow', {
                  label: strings.PropertyPaneLabeldisablePanZoomWindow,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelAbout, // About
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
