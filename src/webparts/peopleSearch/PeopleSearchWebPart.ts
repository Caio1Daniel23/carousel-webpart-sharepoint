import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PeopleSearch from './components/PeopleSearch';
import { IPeopleSearchWebPartProps } from './IPeopleSearchWebPartProps';

export default class PeopleSearchWebPart extends BaseClientSideWebPart<IPeopleSearchWebPartProps> {
  public render(): void {
    const heightValue = Number(this.properties.height);

    const element: React.ReactElement = React.createElement(PeopleSearch, {
      context: this.context,
      placeholderText: this.properties.placeholderText,
      height: !isNaN(heightValue) && heightValue > 0 ? heightValue : 300
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Busca pessoas da organização (Microsoft Graph) com foto e informações de contato.'
          },
          groups: [
            {
              groupName: 'Configurações',
              groupFields: [
                PropertyPaneTextField('placeholderText', {
                  label: 'Texto de exemplo no campo de busca'
                }),
                PropertyPaneTextField('height', {
                  label: 'Altura mínima (px)',
                  description: 'Digite qualquer valor em pixels (ex: 250, 300, 400).',
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
