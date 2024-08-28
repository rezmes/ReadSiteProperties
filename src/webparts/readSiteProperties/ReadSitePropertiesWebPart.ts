import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from "@microsoft/sp-webpart-base";
import { escape, truncate } from "@microsoft/sp-lodash-subset";

import styles from "./ReadSitePropertiesWebPart.module.scss";
import * as strings from "ReadSitePropertiesWebPartStrings";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IReadSitePropertiesWebPartProps {
description: string;

  productname: string;
  productdescription: string;
  productcost: number;
  quantity: number;
  billamount: number;
  discount: number;
  netbillamount: number;
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
    return this.context.spHttpClient
      .get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getAndRenderLists(): void {
    if (Environment.type === EnvironmentType.Local) {
    } else if (
      Environment.type == EnvironmentType.SharePoint ||
      Environment.type == EnvironmentType.ClassicSharePoint
    ) {
      this._getListsOfLists().then((response) => {
        this._renderListOfLists(response.value);
      });
    }
  }
  private _renderListOfLists(items: ISharePointList[]): void {
    let html: string = "";
    items.forEach((item: ISharePointList) => {
      html += `
      <ul class="${styles.list}">
        <li class="${styles.listItem}"><span class="ms-font-1">${item.Title}</span></li>
        <li class="${styles.listItem}"><span class="ms-font-1">${item.Id}</span></li>
      </ul>
    `;
    });
    const listsContainer: Element =
      this.domElement.querySelector("#spListContainer");
    listsContainer.innerHTML = html;
  }





protected onInit(): Promise<void> {
  return new Promise<void>((resolve, _reject) => {
   this.properties.productname = "mouse";
   this.properties.productdescription = "Mouse Description"
  this.properties.quantity = 500;
  this.properties.productcost = 500;

  resolve(undefined);

  })

}

protected get disableReactivePropertyChanges() : boolean {
  return false;
}





  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.readSiteProperties}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">




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

        <table>
          <tr>
            <td><strong>Product Name</strong></td>
            <td><strong>${this.properties.productname}</td>
            </tr>
            <tr>
            <td><strong>Description</strong></td>
            <td>${this.properties.productdescription}</td>
            </tr>
            <tr>
            <td><strong>Product Cost</strong></td>
            <td>${this.properties.productcost}</td>
            </tr>
            <td><strong>Product Quantity</strong></td>
            <td>${this.properties.quantity}</td>

            <tr>
            <td>Bill Amount</td>
            <td>${(this.properties.billamount =
              this.properties.productcost * this.properties.quantity)}</td>
            </tr>
            <tr>
            <td>Discount</td>
            <td>${(this.properties.discount =
              this.properties.billamount * 0.1)}</td>
            </tr>
            <tr>
            <td>Net Bill Amount</td>
            <td>${(this.properties.netbillamount =
              this.properties.billamount - this.properties.discount)}</td>
            </tr>
            </table>



            </div>
          </div>
        </div>
        <div Id="spListContainer"/>
      </div>`;
    this._getAndRenderLists();

  }

  // protected get dataVersion(): Version {
  //   return Version.parse("1.0");
  // }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription,
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField("description", {
  //                 label: strings.DescriptionFieldLabel,
  //               }),
  //             ],
  //           },
  //         ],
  //       },
  //     ],
  //   };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: "Product Details",
              groupFields: [
                PropertyPaneTextField("productname", {
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "please enter product name", "description": "Name property filed",
                }),
                PropertyPaneTextField("productdescription", {
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "please enter product Description",  "description": "description property filed",
                }),
                PropertyPaneTextField("productcost", {
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "please enter product Cost",  "description": "Number property filed",
                }),
                PropertyPaneTextField("quantity", {
                  label: "Product Quantity",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 5000,
                  placeholder: "please enter product Quantity", "description": "Number property filed",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
