import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'NoticiasWebPartStrings';
import Noticias from './components/Noticias';
import { INoticiasProps, LayoutMode } from './components/INoticiasProps';

export interface INoticiasWebPartProps {
  title: string;
  layoutMode: LayoutMode;
  pageSize: number;
  carouselCount: number;
  showAllNews: boolean;
  compactMode: boolean;
  showAuthor: boolean;
  showViews: boolean;
  showDate: boolean;
}

export default class NoticiasWebPart extends BaseClientSideWebPart<INoticiasWebPartProps> {
  protected onInit(): Promise<void> {
    // Defaults for a brand-new instance of the web part.
    if (this.properties.title === undefined) this.properties.title = 'Notícias';
    if (this.properties.layoutMode === undefined) this.properties.layoutMode = 'list';
    if (this.properties.pageSize === undefined) this.properties.pageSize = 10;
    if (this.properties.carouselCount === undefined) this.properties.carouselCount = 5;
    if (this.properties.showAllNews === undefined) this.properties.showAllNews = false;
    if (this.properties.compactMode === undefined) this.properties.compactMode = false;
    if (this.properties.showAuthor === undefined) this.properties.showAuthor = false;
    if (this.properties.showViews === undefined) this.properties.showViews = false;
    if (this.properties.showDate === undefined) this.properties.showDate = true;
    return Promise.resolve();
  }

  public render(): void {
    const element: React.ReactElement<INoticiasProps> = React.createElement(Noticias, {
      context: this.context,
      title: this.properties.title,
      layoutMode: this.properties.layoutMode,
      pageSize: this.properties.pageSize,
      carouselCount: this.properties.carouselCount,
      showAllNews: this.properties.showAllNews,
      compactMode: this.properties.compactMode,
      showAuthor: this.properties.showAuthor,
      showViews: this.properties.showViews,
      showDate: this.properties.showDate
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Refresh the property pane when the layout mode changes, so the
  // "itens no carrossel" slider only shows up for the carousel layout.
  protected onPropertyPaneFieldChanged(propertyPath: string): void {
    if (propertyPath === 'layoutMode') {
      this.context.propertyPane.refresh();
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const isCarousel = this.properties.layoutMode === 'carousel';

    const displayFields = [
      PropertyPaneChoiceGroup('layoutMode', {
        label: strings.LayoutModeFieldLabel,
        options: [
          { key: 'list', text: strings.LayoutModeListLabel },
          { key: 'carousel', text: strings.LayoutModeCarouselLabel }
        ]
      }),
      ...(isCarousel
        ? [
            PropertyPaneSlider('carouselCount', {
              label: strings.CarouselCountFieldLabel,
              min: 3,
              max: 8,
              step: 1,
              value: this.properties.carouselCount
            })
          ]
        : []),
      PropertyPaneSlider('pageSize', {
        label: strings.PageSizeFieldLabel,
        min: 5,
        max: 15,
        step: 5,
        value: this.properties.pageSize,
        disabled: this.properties.showAllNews
      }),
      PropertyPaneToggle('showAllNews', {
        label: strings.ShowAllNewsFieldLabel
      })
    ];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('title', { label: strings.TitleFieldLabel }),
                ...displayFields
              ]
            },
            {
              groupName: strings.DisplayGroupName,
              groupFields: [
                PropertyPaneToggle('compactMode', { label: strings.CompactModeFieldLabel }),
                PropertyPaneToggle('showDate', { label: strings.ShowDateFieldLabel }),
                PropertyPaneToggle('showAuthor', { label: strings.ShowAuthorFieldLabel }),
                PropertyPaneToggle('showViews', { label: strings.ShowViewsFieldLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
