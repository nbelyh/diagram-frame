import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';

import { PropertyPaneSizeFieldComponent } from './PropertyPaneSizeFieldComponent';

export function PropertyPaneSizeField(targetProperty: string, props: {
  value: string;
  label: string;
  description: string;
  screenUnits: string;
}): IPropertyPaneField<IPropertyPaneCustomFieldProps> {

  return {
    targetProperty: targetProperty,
    type: PropertyPaneFieldType.Custom,
    properties: {
      key: targetProperty,

      onRender: (parent: HTMLElement, context: any, changeCallback: (targetProperty: string, newValue: any) => void) => {
        return ReactDom.render(
          <PropertyPaneSizeFieldComponent
            value={props.value}
            label={props.label}
            description={props.description}
            screenUnits={props.screenUnits}
            setValue={(val) => changeCallback(targetProperty, val)}
          />, parent);
      },

      onDispose(parent: HTMLElement): void {
        ReactDom.unmountComponentAtNode(parent);
      }
    }
  };
}
