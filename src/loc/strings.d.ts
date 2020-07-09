declare interface IWebPartStrings {
  BasicGroupName: string;
  UrlFieldLabel: string;
  DocumentPickerTitle: string;
}

declare module 'WebPartStrings' {
  const strings: IWebPartStrings;
  export = strings;
}
