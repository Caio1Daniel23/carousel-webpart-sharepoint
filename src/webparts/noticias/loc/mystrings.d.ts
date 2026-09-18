declare interface INoticiasWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  LayoutModeFieldLabel: string;
  LayoutModeListLabel: string;
  LayoutModeCarouselLabel: string;
  PageSizeFieldLabel: string;
  CarouselCountFieldLabel: string;
  ShowAllNewsFieldLabel: string;
  DisplayGroupName: string;
  CompactModeFieldLabel: string;
  ShowAuthorFieldLabel: string;
  ShowViewsFieldLabel: string;
  ShowDateFieldLabel: string;
}

declare module 'NoticiasWebPartStrings' {
  const strings: INoticiasWebPartStrings;
  export = strings;
}
