import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "PnpReactHooksExamplesWebPartStrings";
import PnpReactHooksExamples from "./components/PnpReactHooksExamples";
import { IPnpReactHooksExamplesProps } from "./components/IPnpReactHooksExamplesProps";

// Import pnp-react-hooks and pnpjs
import {
  createProviderElement,
  PnpHookGlobalOptions,
  FetchWithAbort,
} from "pnp-react-hooks";
import { spfi, SPFx } from "@pnp/sp";

export interface IPnpReactHooksExamplesWebPartProps {
  description: string;
}

export default class PnpReactHooksExamplesWebPart extends BaseClientSideWebPart<IPnpReactHooksExamplesWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _hookOptions: PnpHookGlobalOptions;

  public render(): void {
    const child: React.ReactElement<IPnpReactHooksExamplesProps> =
      React.createElement(PnpReactHooksExamples, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        options: this._hookOptions,
      });

    // Create provider as root element
    const element = createProviderElement(this._hookOptions, child);

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then((message) => {
      const sp = spfi().using(SPFx(this.context), FetchWithAbort());
      this._environmentMessage = message;

      // Setup your default PnP React hooks options.
      this._hookOptions = {
        sp: sp,
        disabled: "auto",
        error: console.error,
        keepPreviousState: false,
      };
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
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment,
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
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
