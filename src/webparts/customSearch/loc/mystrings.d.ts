declare interface ICustomSearchStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'customSearchStrings' {
  const strings: ICustomSearchStrings;
  export = strings;
}
