import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "ConsumerWebPartStrings";
import Costumer from "./components/Consumer";
import { ICostumerProps } from "./components/IConsumerProps";
import { DynamicProperty } from "@microsoft/sp-component-base";

import {
  DynamicDataSharedDepth,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField,
  IWebPartPropertiesMetadata,
} from "@microsoft/sp-webpart-base";

export interface IConsumerWebPartProps {
  description: string;
  DeptTitleId: DynamicProperty<string>;
}

export default class ConsumerWebPart extends BaseClientSideWebPart<IConsumerWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    const element: React.ReactElement<ICostumerProps> = React.createElement(
      Costumer,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        context: this.context,
        DeptTitleId: this.properties.DeptTitleId,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error("Unknown host");
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      DeptTitleId: { dynamicPropertyType: "string" },
    };
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
              groupFields: [
                PropertyPaneDynamicFieldSet({
                  label: " Select Department ID",
                  fields: [
                    PropertyPaneDynamicField("DeptTitleId", {
                      label: "Department ID",
                    }),
                  ],
                  sharedConfiguration: {
                    depth: DynamicDataSharedDepth.Property,
                    source: {
                      sourcesLabel:
                        " Select the web part containing the list of Departments",
                    },
                  },
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
