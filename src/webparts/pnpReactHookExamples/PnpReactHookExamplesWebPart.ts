import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import
{
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PnpReactHookExamplesWebPartStrings';
import PnpReactHookExamples from './components/PnpReactHookExamples';
import { IPnpReactHookExamplesProps } from './components/IPnpReactHookExamplesProps';
import { PnpHookGlobalOptions } from 'pnp-react-hooks';
import { spfi, SPFx} from '@pnp/sp';

export interface IPnpReactHookExamplesWebPartProps
{
  description: string;
}



export default class PnpReactHookExamplesWebPart extends BaseClientSideWebPart<IPnpReactHookExamplesWebPartProps>
{

  private _hookOptions: PnpHookGlobalOptions;

  protected onInit(): Promise<void>
  {
    return super.onInit().then(async (_) =>
    {
      const sp = spfi().using(SPFx(this.context));

      this._hookOptions = {
        disabled: "auto",
        keepPreviousState: false,
        error: console.error,
        sp: sp
      };
    });
  }

  public render(): void
  {
    const element: React.ReactElement<IPnpReactHookExamplesProps> = React.createElement(
      PnpReactHookExamples,
      {
        description: this.properties.description,
        options: this._hookOptions
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void
  {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version
  {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration
  {
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
