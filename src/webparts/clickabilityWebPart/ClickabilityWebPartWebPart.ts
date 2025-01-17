import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import { escape } from '@microsoft/sp-lodash-subset';

import { sp } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";

import styles from './ClickabilityWebPartWebPart.module.scss';
import * as strings from 'ClickabilityWebPartWebPartStrings';

export interface IClickabilityWebPartWebPartProps {
  description: string;
}

export default class ClickabilityWebPartWebPart extends BaseClientSideWebPart<IClickabilityWebPartWebPartProps> {

  protected onInit(): Promise<void> {

    console.log('onInit executed');

    
    return super.onInit().then( _ => {

      sp.setup({
        spfxContext: this.context
      });
  
      console.log("spfxContext: this.contxt is", this.context)


    })


    
  }

  public  render(): void {

    this.domElement.innerHTML = "";

    sp.web.select("Title").get().then( (data) => {

      console.log("web return is ",data);

      this.domElement.innerHTML += `Web site title is ${data.Title}`

    })

    this.getListNames();

    
  }

  private async getListNames()  {

    let lists = await sp.web.lists();
    

    console.log('lists',lists)

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
