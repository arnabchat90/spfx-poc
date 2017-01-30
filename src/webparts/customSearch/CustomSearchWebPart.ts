import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'customSearchStrings';
import CustomSearch from './components/CustomSearch';
import { ICustomSearchProps } from './components/ICustomSearchProps';
import { ICustomSearchWebPartProps } from './ICustomSearchWebPartProps';

import SearchPanel from './components/SearchPanel';
import {ISearchPanelProps} from './components/ISearchPanelProps';


export default class CustomSearchWebPart extends BaseClientSideWebPart<ICustomSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISearchPanelProps> = React.createElement(
      SearchPanel,
      {
        siteUrl: "https://nvsdev.sharepoint.com/sites/spfx-dev",
        httpClient : this.context.spHttpClient,
        description : this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
