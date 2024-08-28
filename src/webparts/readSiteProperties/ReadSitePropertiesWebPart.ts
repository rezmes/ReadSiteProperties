import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./ReadSitePropertiesWebPart.module.scss";
import * as strings from "ReadSitePropertiesWebPartStrings";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IReadSitePropertiesWebPartProps {
  description: string;
  environmenttitle: string;
}

export interface ISharePointList {
  Title: string;
  Id: string;
}

export interface ISharePointLists {
value: ISharePointList[];
}



export default class ReadSitePropertiesWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartProps> {

// filter=Hidden eq false instead of select=Title,Id
private _getListsOfLists(): Promise<ISharePointLists> {
  return this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
      SPHttpClient.configurations.v1
    ).then((response: SPHttpClientResponse) => {
      // console.log("Response::::::::",response.json());
      return response.json();

    });

  }


  private _getAndRenderLists(): void {
    if(Environment.type ===EnvironmentType.Local){

    }
    else if (
      Environment.type == EnvironmentType.SharePoint ||
       Environment.type == EnvironmentType.ClassicSharePoint
      ){
      this._getListsOfLists()
        .then((response)=>{
           this._renderListOfLists(response.value)
      });
    }
  }
private _renderListOfLists(items: ISharePointList[]):void {
  let html: string = "";
  items.forEach((item: ISharePointList) => {
    html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}"><span class="ms-font-1">${item.Title}</span></li>
        <li class="${styles.listItem}"><span class="ms-font-1">${item.Id}</span></li>
      </ul>
    `;
  });
  const listsContainer: Element = this.domElement.querySelector(
    "#spListContainer"
  );
  listsContainer.innerHTML = html;
}








  private _findOutEnvironment(): void {
    //Local environment
    if (Environment.type === EnvironmentType.Local) {
      this.properties.environmenttitle = "Local SharePoint Environment";
    } else if (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    ) {
      this.properties.environmenttitle = "Online SharePoint Environment";
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.readSiteProperties}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${
                styles.subTitle
              }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>

              <p class="${styles.description}">absolute URL ${escape(
      this.context.pageContext.web.absoluteUrl
    )}</p>
              <p class="${styles.description}">Title ${escape(
      this.context.pageContext.web.title
    )}</p>
              <p class="${styles.description}">Relative URL ${escape(
      this.context.pageContext.web.serverRelativeUrl
    )}</p>
              <p class="${styles.description}">User Name ${escape(
      this.context.pageContext.user.displayName
    )}</p>

              <p class="${styles.description}">Environment ${
      Environment.type
    }</p>
              <p class="${styles.description}">Environment ${escape(
      this.properties.environmenttitle
    )}</p>

      <ul>
        <li><strong>curent Culture Name</strong>: ${escape(
          this.context.pageContext.cultureInfo.currentCultureName
        )}</li>
        <li><strong>curent Culture UI Name</strong>: ${escape(
          this.context.pageContext.cultureInfo.currentUICultureName
        )}</li>
        <li><strong>is right to left?</strong>: ${
          this.context.pageContext.cultureInfo.isRightToLeft
        }</li>

      </ul>






              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <div Id="spListContainer"/>
      </div>`;
        this._getAndRenderLists();
    this._findOutEnvironment();
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
