declare interface IWebPartStrings {
  FieldVisioFile: string;
  FieldVisioFileBrowse: string;
  FieldZoom: string;
  BasicGroupName: string;
  FieldWidth: string;
  FieldHeight: string;
  View: string;
  FieldStartPage: string;
}

declare module 'WebPartStrings' {
  const strings: IWebPartStrings;
  export = strings;
}
