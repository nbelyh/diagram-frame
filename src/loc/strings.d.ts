declare interface IWebPartStrings {
  FieldVisioFile: string;
  FieldVisioFileBrowse: string;
  FieldZoom: string;
  BasicGroupName: string;
  FieldWidth: string;
  FieldHeight: string;
  View: string;
  FieldStartPage: string;
  PropertyPaneLabelDrawingDisplay: string;
  PropertyPaneLabelAppearance: string;
  PropertyPaneLabelInteractivity: string;
  PropertyPaneLabelAbout: string;
  PropertyPaneLabelhideToolbars: string;
  PropertyPaneLabelhideDiagramBoundary: string;
  PropertyPaneLabelhideBorders: string;
  PropertyPaneLabeldisableHyperlinks: string;
  PropertyPaneLabeldisablePan: string;
  PropertyPaneLabeldisableZoom: string;
  PropertyPaneLabeldisablePanZoomWindow: string;
  FieldHeightDescription: string;
  FieldWidthDescription: string;
  FieldStartPageDescription: string;
  FieldZoomDescription: string;
}

declare module 'WebPartStrings' {
  const strings: IWebPartStrings;
  export = strings;
}
