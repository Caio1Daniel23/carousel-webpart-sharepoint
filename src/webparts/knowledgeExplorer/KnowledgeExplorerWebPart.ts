import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import KnowledgeExplorer from './components/KnowledgeExplorer';
import { IKnowledgeExplorerWebPartProps, ICustomFolderImage } from './IKnowledgeExplorerWebPartProps';

import { PropertyFieldFolderPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFolderPicker';
import { PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { FilePicker } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/propertyFields/filePicker/filePickerControls/FilePicker.types';

export default class KnowledgeExplorerWebPart extends BaseClientSideWebPart<IKnowledgeExplorerWebPartProps> {
  public render(): void {
    const heightValue = Number(this.properties.height);

    // O seletor de pasta pode gravar o objeto {Name, ServerRelativeUrl} inteiro em vez de
    // só o texto do caminho, dependendo da versão do controle — aceitamos os dois formatos.
    const rawRootPath: any = this.properties.rootFolderPath as any;
    const normalizedRootPath: string =
      typeof rawRootPath === 'string' ? rawRootPath : rawRootPath?.ServerRelativeUrl || '';

    const element: React.ReactElement = React.createElement(KnowledgeExplorer, {
      context: this.context,
      rootFolderPath: normalizedRootPath,
      rootTitle: this.properties.rootTitle,
      height: !isNaN(heightValue) && heightValue > 0 ? heightValue : 500,
      defaultCardImage: this.properties.defaultCardImage,
      customFolderImages: this.properties.customFolderImages || [],
      displayMode: this.displayMode,
      uploadPickedImage: this.onImageSave.bind(this),
      onSetFolderImage: this.onSetFolderImage.bind(this),
      onRemoveFolderImage: this.onRemoveFolderImage.bind(this)
    });

    ReactDom.render(element, this.domElement);
  }

  // Reagir a alternância entre modo de edição/leitura (ex: ao clicar "Editar página" ou
  // "Publicar/Sair da edição") — sem isso, o componente React não saberia que precisa
  // mostrar ou esconder os "3 pontinhos" dos cards até a próxima ação que force um render.
  protected onDisplayModeChanged(oldDisplayMode: DisplayMode): void {
    this.render();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // Grava um arquivo escolhido/enviado do computador na biblioteca Site Assets do site atual.
  // (mesma lógica usada no carrossel para upload de imagens)
  private async onImageSave(filePickerResult: IFilePickerResult): Promise<string> {
    if (filePickerResult.fileAbsoluteUrl) {
      return filePickerResult.fileAbsoluteUrl;
    }

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
    return json.ServerRelativeUrl ? `${window.location.origin}${json.ServerRelativeUrl}` : '';
  }

  // Grava (ou substitui) a imagem específica de uma pasta na propriedade da web part.
  // Chamado direto pelo "3 pontinhos" do card, sem precisar abrir o painel de propriedades —
  // por isso também atualizamos o painel (caso já esteja aberto) pra manter a lista sincronizada.
  private onSetFolderImage(folderPath: string, imageUrl: string): void {
    const current: ICustomFolderImage[] = this.properties.customFolderImages || [];
    const withoutThisFolder = current.filter((c) => c.folderPath !== folderPath);
    this.properties.customFolderImages = [...withoutThisFolder, { folderPath, imageUrl }];
    this.render();
    this.context.propertyPane.refresh();
  }

  // Remove a imagem específica de uma pasta — ela volta a usar a imagem padrão.
  private onRemoveFolderImage(folderPath: string): void {
    const current: ICustomFolderImage[] = this.properties.customFolderImages || [];
    this.properties.customFolderImages = current.filter((c) => c.folderPath !== folderPath);
    this.render();
    this.context.propertyPane.refresh();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description:
              'Aponte para a pasta raiz de uma biblioteca de documentos. As subpastas viram categorias ' +
              'automaticamente, com a contagem de arquivos, e a navegação acontece dentro da própria web part.'
          },
          groups: [
            {
              groupName: 'Configurações',
              groupFields: [
                PropertyFieldFolderPicker('rootFolderPath', {
                  context: this.context as any,
                  label: 'Pasta raiz',
                  rootFolder: {
                    Name: this.context.pageContext.web.title,
                    ServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl
                  },
                  defaultFolder: {
                    Name: this.context.pageContext.web.title,
                    ServerRelativeUrl: this.properties.rootFolderPath || this.context.pageContext.web.serverRelativeUrl
                  },
                  selectedFolder: {
                    Name: '',
                    ServerRelativeUrl: this.properties.rootFolderPath || this.context.pageContext.web.serverRelativeUrl
                  },
                  canCreateFolders: false,
                  onSelect: (folder) => {
                    this.properties.rootFolderPath = folder.ServerRelativeUrl;
                    this.render();
                  },
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  key: 'rootFolderPickerFieldId'
                } as any),
                PropertyPaneTextField('rootTitle', {
                  label: 'Nome exibido no início da navegação (ex: Base de Conhecimento)'
                }),
                PropertyPaneTextField('height', {
                  label: 'Altura mínima (px)',
                  description: 'Digite qualquer valor em pixels (ex: 400, 500, 600).',
                  onGetErrorMessage: (value: string) => {
                    const num = Number(value);
                    if (value === '' || isNaN(num) || num <= 0) {
                      return 'Digite um número de pixels maior que 0.';
                    }
                    return '';
                  }
                })
              ]
            },
            {
              groupName: 'Imagens dos cards (1º e 2º nível de pastas)',
              groupFields: [
                PropertyFieldFilePicker('defaultCardImage', {
                  context: this.context as any,
                  key: 'defaultCardImageFieldId',
                  label: 'Imagem padrão (aplica em todos os cards de 1º e 2º nível)',
                  buttonLabel: this.properties.defaultCardImage ? 'Alterar imagem' : 'Escolher imagem',
                  accepts: ['.gif', '.jpg', '.jpeg', '.png', '.webp', '.svg'],
                  onSave: async (result: IFilePickerResult) => {
                    this.properties.defaultCardImage = await this.onImageSave(result);
                    this.render();
                  }
                } as any),
                PropertyFieldCollectionData('customFolderImages', {
                  key: 'customFolderImages',
                  label: 'Imagens específicas por pasta (opcional, sobrescreve a padrão)',
                  panelHeader: 'Editar imagens por pasta',
                  manageBtnLabel: 'Gerenciar imagens por pasta',
                  value: this.properties.customFolderImages,
                  fields: [
                    {
                      id: 'folderPath',
                      title: 'Caminho da pasta (ex: /sites/intranet/BaseConhecimento/Contencioso)',
                      type: CustomCollectionFieldType.string
                    },
                    {
                      id: 'imageUrl',
                      title: 'Imagem',
                      type: CustomCollectionFieldType.custom,
                      onCustomRender: (field, value, onUpdate) => {
                        return React.createElement(
                          'div',
                          { style: { display: 'flex', flexDirection: 'column', alignItems: 'flex-start', gap: 6, minWidth: 140 } },
                          value
                            ? React.createElement('img', {
                                src: value,
                                style: { width: 130, height: 70, objectFit: 'cover', borderRadius: 2, border: '1px solid #edebe9' }
                              })
                            : null,
                          React.createElement(FilePicker, {
                            context: this.context as any,
                            buttonIcon: 'Photo2',
                            buttonLabel: value ? 'Alterar' : 'Escolher imagem',
                            accepts: ['.gif', '.jpg', '.jpeg', '.png', '.webp', '.svg'],
                            onSave: async (result: IFilePickerResult) => {
                              const url = await this.onImageSave(result);
                              onUpdate(field.id, url);
                            }
                          } as any)
                        );
                      }
                    }
                  ],
                  disabled: false
                } as any)
              ]
            }
          ]
        }
      ]
    };
  }
}
