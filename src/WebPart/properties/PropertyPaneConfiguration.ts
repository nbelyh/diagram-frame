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
                PropertyPaneVersionField(context.manifest.version)
              ]
            }
          ]
        }
      ]
    };
  }
}
