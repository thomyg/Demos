import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloThriveWebPart.module.scss';
import * as strings from 'HelloThriveWebPartStrings';

export interface IHelloThriveWebPartProps {
  description: string;
}

export default class HelloThriveWebPart extends BaseClientSideWebPart<IHelloThriveWebPartProps> {

  private _teamsContext: microsoftTeams.Context;

  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  public render(): void {

    let title: string = '';
    let subTitle: string = '';
    let siteTabTitle: string = '';
    let imgUrl: string = '';

    if (this._teamsContext) {
      // We have teams context for the web part
      title = "Welcome to Teams Thrive Conference!";
      subTitle = "Building custom enterprise tabs for your business.";
      siteTabTitle = "We are in the context of following Team: " + this._teamsContext.teamName;
      imgUrl = "https://media-cdn.tripadvisor.com/media/photo-s/0c/c4/58/33/ljubljana-by-night.jpg";
    }
    else
    {
      // We are rendered in normal SharePoint context
      title = "Welcome to SharePoint Thrive Conference!";
      subTitle = "Customize SharePoint experiences using Web Parts.";
      siteTabTitle = "We are in the context of following site: " + this.context.pageContext.web.title;
      imgUrl = "http://www.sloveniaguides.si/media/rokgallery/b/b2799bf5-7d2e-4c16-82bd-d3df7d7bc68d/day-ljubljana-02.jpg";
    }

    this.domElement.innerHTML = `
      <div class="${ styles.helloThrive }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">${title}</span>
              <p class="${ styles.subTitle }">${subTitle}</p>
              <p class="${ styles.description }">${siteTabTitle}</p>
              <p class="${ styles.description }">Description property value - ${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
            <div style="width:200px;heigh:200px">
              <img style="width:300px;" src="${imgUrl}"/>
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
