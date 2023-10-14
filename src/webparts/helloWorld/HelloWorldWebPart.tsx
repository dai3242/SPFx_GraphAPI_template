import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "HelloWorldWebPartStrings";
import HelloWorld from "./components/HelloWorld";
import { IHelloWorldProps } from "./components/IHelloWorldProps";

// "IHelloWorldProps" is the type for properties in property pane.
export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldProps> {
  public render(): void {
    ReactDom.render(
      // context is for executing graphAPI, properties are for property pane
      <HelloWorld context={this.context} properties={this.properties} />,
      this.domElement
    );
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
                  label: "description",
                }),
                PropertyPaneTextField("siteName", {
                  label: "siteName",
                }),
                PropertyPaneTextField("test", {
                  label: "This is mulit-line text field",
                  multiline: true,
                }),
                PropertyPaneCheckbox("test1", {
                  text: "This is checkbox",
                }),
                PropertyPaneDropdown("test2", {
                  label: "This is dropdown",
                  options: [
                    { key: "1", text: "One" },
                    { key: "2", text: "Two" },
                    { key: "3", text: "Three" },
                    { key: "4", text: "Four" },
                  ],
                }),
                PropertyPaneToggle("test3", {
                  label: "This is toggle",
                  onText: "On",
                  offText: "Off",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
