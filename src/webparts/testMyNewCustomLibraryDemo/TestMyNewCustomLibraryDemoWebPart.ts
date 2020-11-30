import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestMyNewCustomLibraryDemoWebPart.module.scss';
import * as strings from 'TestMyNewCustomLibraryDemoWebPartStrings';

import * as myLibrary from 'my-new-custom-library';

export interface ITestMyNewCustomLibraryDemoWebPartProps {
  description: string;
}

export default class TestMyNewCustomLibraryDemoWebPart extends BaseClientSideWebPart<ITestMyNewCustomLibraryDemoWebPartProps> {

  public render(): void {

    const myInstance = new myLibrary.MyNewCustomLibraryDemoLibrary();

    this.domElement.innerHTML = `
      <div class="${ styles.testMyNewCustomLibraryDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>

              <p>Calling libarary function</p>
              <p>${myInstance.getCurrentTime()}</p>

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
