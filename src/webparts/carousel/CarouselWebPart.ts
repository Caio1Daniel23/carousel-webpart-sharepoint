import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'CarouselWebPartStrings';
import Carousel from './components/Carousel';
import { ICarouselWebPartProps, ISlide, CtaType } from './ICarouselWebPartProps';

import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { FilePicker } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker.types';

export default class CarouselWebPart extends BaseClientSideWebPart<ICarouselWebPartProps> {
  public render(): void {
    const element: React.ReactElement = React.createElement(Carousel, {
      slides: this.properties.slides || [],
      autoplay: this.properties.autoplay,
      transitionTime: this.properties.transitionTime,
      height: this.properties.height,
      showArrows: this.properties.showArrows,
      showDots: this.properties.showDots
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Callback usado pelo seletor de imagem (upload local ou biblioteca do SharePoint) de cada slide.
  private async onImageSave(filePickerResult: IFilePickerResult, currentItem: ISlide): Promise<string> {
    // Quando o item já veio de uma URL/biblioteca, o próprio resultado traz o endereço.
    if (filePickerResult.fileAbsoluteUrl) {
      return filePickerResult.fileAbsoluteUrl;
    }

    // Quando o usuário faz upload de um arquivo do computador, gravamos em Site Assets.
    const fileContent = await filePickerResult.downloadFileContent();
    const webServerRelativeUrl = this.context.pageContext.web.serverRelativeUrl.replace(/\/$/, '');
    const libraryServerRelativeUrl = `${webServerRelativeUrl}/SiteAssets`;

    const response = await this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryServerRelativeUrl}')/Files/add(url='${filePickerResult.fileName}',overwrite=true)`,
      SPHttpClient.configurations.v1,
      {
        headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/octet-stream' },
        body: fileContent
      }
    );

    const json = await response.json();
    return json.ServerRelativeUrl
      ? `${window.location.origin}${json.ServerRelativeUrl}`
      : '';
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.SlidesFieldLabel,
              groupFields: [
                PropertyFieldCollectionData('slides', {
                  key: 'slides',
                  label: 'Slides do carrossel',
                  panelHeader: 'Editar slides',
                  manageBtnLabel: 'Gerenciar slides',
                  value: this.properties.slides,
                  fields: [
                    {
                      id: 'imageUrl',
                      title: 'Imagem',
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item: ISlide) => {
                        return React.createElement(
                          'div',
                          { style: { display: 'flex', alignItems: 'center', gap: 8 } },
                          value
                            ? React.createElement('img', {
                                src: value,
                                style: { width: 60, height: 40, objectFit: 'cover', borderRadius: 2 }
                              })
                            : null,
                          React.createElement(FilePicker, {
                            context: this.context as any,
                            buttonIcon: 'Photo2',
                            buttonLabel: value ? 'Alterar imagem' : 'Escolher imagem',
                            accepts: ['.gif', '.jpg', '.jpeg', '.png', '.webp', '.svg'],
                            hideRecentTab: false,
                            hideStockImages: true,
                            hideWebSearchTab: false,
                            hideOrganisationalAssetTab: false,
                            hideSiteFilesTab: false,
                            hideLinkUploadTab: false,
                            hideLocalUploadTab: false,
                            onSave: async (result: IFilePickerResult) => {
                              const url = await this.onImageSave(result, item);
                              onUpdate(field.id, url);
                            },
                            onChanged: undefined
                          } as any)
                        );
                      }
                    },
                    {
                      id: 'imageAlt',
                      title: 'Texto alternativo da imagem',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'preHeader',
                      title: 'Pré-cabeçalho',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'title',
                      title: 'Título',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'description',
                      title: 'Descrição',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'ctaType',
                      title: 'Tipo de chamada à ação',
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        { key: 'none' as CtaType, text: 'Nenhuma' },
                        { key: 'button' as CtaType, text: 'Botão' },
                        { key: 'icon' as CtaType, text: 'Ícone' },
                        { key: 'text' as CtaType, text: 'Texto' },
                        { key: 'card' as CtaType, text: 'Cartão inteiro' }
                      ]
                    },
                    {
                      id: 'ctaText',
                      title: 'Texto da chamada à ação',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'ctaIcon',
                      title: 'Ícone (nome Fluent UI, ex: ChevronRight)',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'ctaLink',
                      title: 'Link',
                      type: CustomCollectionFieldType.url
                    }
                  ],
                  disabled: false
                } as any)
              ]
            },
            {
              groupName: strings.BehaviorGroupName,
              groupFields: [
                PropertyPaneToggle('autoplay', {
                  label: 'Reprodução automática'
                }),
                PropertyPaneSlider('transitionTime', {
                  label: 'Tempo entre slides (segundos)',
                  min: 2,
                  max: 30,
                  step: 1,
                  showValue: true
                }),
                PropertyPaneToggle('showArrows', {
                  label: 'Exibir setas de navegação'
                }),
                PropertyPaneToggle('showDots', {
                  label: 'Exibir indicadores (bolinhas)'
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneSlider('height', {
                  label: 'Altura do carrossel (px)',
                  min: 150,
                  max: 900,
                  step: 10,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
