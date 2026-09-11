import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneDropdown
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
    const heightValue = Number(this.properties.height);
    const element: React.ReactElement = React.createElement(Carousel, {
      slides: this.properties.slides || [],
      autoplay: this.properties.autoplay,
      transitionTime: this.properties.transitionTime,
      height: !isNaN(heightValue) && heightValue > 0 ? heightValue : 400,
      showArrows: this.properties.showArrows,
      showDots: this.properties.showDots,
      dotsPosition: this.properties.dotsPosition || 'inside'
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
                  enableSorting: true,
                  fields: [
                    {
                      id: 'imageUrl',
                      title: 'Imagem',
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate, item: ISlide) => {
                        return React.createElement(
                          'div',
                          { style: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 6, minWidth: 160 } },
                          value
                            ? React.createElement('img', {
                                src: value,
                                style: {
                                  width: 150,
                                  height: 90,
                                  objectFit: 'cover',
                                  borderRadius: 2,
                                  border: '1px solid #edebe9'
                                }
                              })
                            : React.createElement(
                                'div',
                                {
                                  style: {
                                    width: 150,
                                    height: 90,
                                    borderRadius: 2,
                                    border: '1px dashed #c8c6c4',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    color: '#a19f9d',
                                    fontSize: 11,
                                    textAlign: 'center'
                                  }
                                },
                                'Sem imagem'
                              ),
                          React.createElement(FilePicker, {
                            context: this.context as any,
                            buttonIcon: 'Photo2',
                            buttonLabel: value ? 'Alterar' : 'Escolher imagem',
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
                    },
                    {
                      id: 'ctaOpenNewTab',
                      title: 'Abrir link em nova janela',
                      type: CustomCollectionFieldType.boolean
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
                }),
                PropertyPaneDropdown('dotsPosition', {
                  label: 'Posição das bolinhas',
                  options: [
                    { key: 'inside', text: 'Dentro do carrossel (sobre a imagem)' },
                    { key: 'below', text: 'Abaixo do carrossel (fora da imagem)' }
                  ]
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneTextField('height', {
                  label: 'Altura do carrossel (px)',
                  description: 'Digite qualquer valor em pixels (ex: 245, 400, 512).',
                  onGetErrorMessage: (value: string) => {
                    const num = Number(value);
                    if (value === '' || isNaN(num) || num <= 0) {
                      return 'Digite um número de pixels maior que 0.';
                    }
                    return '';
                  }
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
