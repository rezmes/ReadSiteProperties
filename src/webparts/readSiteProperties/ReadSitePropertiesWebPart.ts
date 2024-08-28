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

export interface IReadSitePropertiesWebPartProps {
  description: string;
  environmenttitle: string;
}

export default class ReadSitePropertiesWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartProps> {
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
      </div>`;

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
