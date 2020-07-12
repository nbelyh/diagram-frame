declare interface IWebPartStrings {
  FieldVisioFile: string;
  FieldVisioFileBrowse: string;
  FieldZoom: string;
  BasicGroupName: string;
  FieldWidth: string;
  FieldHeight: string;
  Toolbars: string;
  FieldShowToolbars: string;
  FieldShowBorders: string;
}

declare module 'WebPartStrings' {
  const strings: IWebPartStrings;
  export = strings;
}
