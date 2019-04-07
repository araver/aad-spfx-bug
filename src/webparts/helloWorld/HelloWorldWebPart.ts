import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { AadHttpClientConfiguration, AadHttpClientFactory, AadHttpClient } from '@microsoft/sp-http';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        Loading
      </div>`;
      this._render();
  }

  private async _render(){
    try{
      let tp = await this.context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
      let response = await tp.get("https://graph.microsoft.com/v1.0/me", AadHttpClient.configurations.v1);
      this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        ${await response.text()}
      </div>`;
    }catch(ex){
      this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        ${ex}
      </div>`;
    }
    
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
