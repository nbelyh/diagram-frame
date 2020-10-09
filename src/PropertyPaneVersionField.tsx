import * as React from 'react';
import * as ReactDom from 'react-dom';
import { IPropertyPaneField, PropertyPaneFieldType, IPropertyPaneCustomFieldProps } from '@microsoft/sp-property-pane';

export function PropertyPaneVersionField(version: string): IPropertyPaneField<IPropertyPaneCustomFieldProps> {

  return {
    targetProperty: '',
    type: PropertyPaneFieldType.Custom,
    properties: {
      key: "version",
      onRender: (parent: HTMLElement, context: any, changeCallback: (targetProperty: string, newValue: any) => void) => {
        const elem = (
          <div>Version: {version}</div>
        );
        return ReactDom.render(elem, parent);
      },

      onDispose(parent: HTMLElement): void {
        ReactDom.unmountComponentAtNode(parent);
      }
    }
  };
}
