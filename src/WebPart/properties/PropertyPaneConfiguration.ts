import * as strings from 'WebPartStrings';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';

import { PropertyPaneVersionField } from './PropertyPaneVersionField';
import { PropertyPaneUrlField } from './PropertyPaneUrlField';
import { PropertyPaneSizeField } from './PropertyPaneSizeField';
import { Defaults } from './Defaults';
import { IWebPartProps } from '../IWebPartProps';

export class PropertyPaneConfiguration {

  public static get(context: WebPartContext, properties: IWebPartProps): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.PropertyPaneLabelDrawingDisplay, // Drawing Display
              groupFields: [
                PropertyPaneUrlField('url', {
                  url: properties.url,
                  context: context,
                  getDefaultFolder: () => Defaults.getDefaultFolder(context),
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
                  inlineLabel: true,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelAppearance, // Appearance
              groupFields: [
                PropertyPaneSizeField('width', {
                  label: strings.FieldWidth,
                  description: strings.FieldWidthDescription,
                  value: properties.width,
                  screenUnits: 'w',
                  getDefaultValue: () => Defaults.getDefaultWidth(context)
                }),

                PropertyPaneSizeField('height', {
                  label: strings.FieldHeight,
                  description: strings.FieldHeightDescription,
                  value: properties.height,
                  screenUnits: 'h',
                  getDefaultValue: () => Defaults.getDefaultHeight(context)
                }),
                PropertyPaneToggle('hideToolbars', {
                  label: strings.PropertyPaneLabelhideToolbars,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('hideDiagramBoundary', {
                  label: strings.PropertyPaneLabelhideDiagramBoundary,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('hideBorders', {
                  label: strings.PropertyPaneLabelhideBorders,
                  inlineLabel: true,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelInteractivity, // Drawing Interactivity
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('disableHyperlinks', {
                  label: strings.PropertyPaneLabeldisableHyperlinks,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('openHyperlinksInNewWindow', {
                  label: "Open Hyperlinks in New Window",
                  disabled: properties.disableHyperlinks,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('forceOpeningOfficeFilesOnline', {
                  label: "Force Open Office Files",
                  disabled: properties.disableHyperlinks,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('disablePan', {
                  label: strings.PropertyPaneLabeldisablePan,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('disableZoom', {
                  label: strings.PropertyPaneLabeldisableZoom,
                  inlineLabel: true,
                }),
                PropertyPaneToggle('disablePanZoomWindow', {
                  label: strings.PropertyPaneLabeldisablePanZoomWindow,
                  inlineLabel: true,
                }),
              ]
            },
            {
              groupName: strings.PropertyPaneLabelAbout, // About
              groupFields: [
                PropertyPaneVersionField(context.manifest.version)
              ]
            }
          ]
        }
      ]
    };
  }
}
