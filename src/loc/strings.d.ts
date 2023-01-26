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
  FieldConfigureLabel: string;
  Error: string;
  Edit: string;
  placeholderIconTextUnableShowVisio: string;
  placeholderIconTextVisioNotSelected: string;
  placeholderIconTextPleaseclickBrowse: string;
  placeholderIconTextPleaseclickSettings: string;
  placeholderIconTextPleaseclickEdit: string;
  placeholderIconTextPleaseclickConfigure: string;
  messageWasTheFileDeleted: string;
  messageArePermissionsMissing: string;
  messageCannotResolveFileURL: string;
  messageSomethingWentWrongResolveURL: string;
  messageVisioFileNotFound: string;
  messageVisioFileCannotAccessed: string;
  percOfScreen: string;
  percOfFrame: string;
  centimeters: string;
  inches: string;
  millimeters: string;
  points: string;
  pixels: string;
  VisioDocument: string;
}

declare module 'WebPartStrings' {
  const strings: IWebPartStrings;
  export = strings;
}
