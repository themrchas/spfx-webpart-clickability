import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp";
import "@pnp/sp/lists";

import styles from './ClickabilityWebPartWebPart.module.scss';
import * as strings from 'ClickabilityWebPartWebPartStrings';

export interface IClickabilityWebPartWebPartProps {
  description: string;
}

export default class ClickabilityWebPartWebPart extends BaseClientSideWebPart<IClickabilityWebPartWebPartProps> {

  protected onInit(): Promise<void> {

    console.log('onInit executed');

    sp.setup({
      spfxContext: this.context
    });

    console.log("spfxContext: this.contxt is", this.context)

    return super.onInit();


    
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.clickabilityWebPart }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
