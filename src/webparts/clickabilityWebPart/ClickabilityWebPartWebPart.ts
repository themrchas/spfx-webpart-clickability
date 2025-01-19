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
import { IListInfo } from '@pnp/sp/lists';

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

      this.domElement.innerHTML += `<h2>Current web is ${data.Title}</h2>`

      this.getListNames().then( (data) =>{
       
        this.domElement.innerHTML += data;

      })


    })

    this.getListNames();
    

    
  }

  private async getListNames(): Promise<string> {

    //Array consisting of Document Library names in the web
    const documentLibraryNames: string[] = [];

    console.log("sp.web is",sp.web);

    let lists = await sp.web.lists();
   // let docLibsRaw: IListInfo = await sp.web.lists.filter("BaseType eq 1").get();
   let docLibsRaw: IListInfo[] = await sp.web.lists.filter("BaseType eq 1").select("Title").get();

   console.log("docLibsRaw has type ",typeof docLibsRaw);
    
    console.log('lists',lists);


    console.log("doc lib information", docLibsRaw);

    let docLibTitles: string = "<h3>The Document Libraries on this site are below</h3>"

    //Create a unordered list of document library titles
    docLibTitles += docLibsRaw.reduce( (acc,item) => {

      return `${acc}<li>${item.Title}</li>`

    },"<h4><ul>") + "</ul></h4>"

    return Promise.resolve(docLibTitles)





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
