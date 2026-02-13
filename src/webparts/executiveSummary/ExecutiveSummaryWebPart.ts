import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ExecutiveSummaryWebPartStrings';
import ExecutiveSummary from './components/ExecutiveSummary';
import { IExecutiveSummaryProps } from './components/IExecutiveSummaryProps';

export interface IExecutiveSummaryWebPartProps {
  description: string;
  listName: string;
  siteUrl: string;
}

export default class ExecutiveSummaryWebPart extends BaseClientSideWebPart<IExecutiveSummaryWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _siteOptions: IPropertyPaneDropdownOption[] = [];
  private _listOptions: IPropertyPaneDropdownOption[] = [];
  private _listsDisabled: boolean = true;

  public render(): void {
    const element: React.ReactElement<IExecutiveSummaryProps> = React.createElement(
      ExecutiveSummary,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        listName: this.properties.listName,
        siteUrl: this.properties.siteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneConfigurationStart(): void {
    this.context.statusRenderer.displayLoadingIndicator(
      this.domElement,
      "Configuration"
    );

    void this._getSiteRootWeb()
      .then((response) => {
        void this._getSites(response["Url"])
          .then((response1) => {
            const sites: IPropertyPaneDropdownOption[] = [];
            sites.push({
              key: this.context.pageContext.web.absoluteUrl,
              text: "This Site",
            });
            sites.push({ key: "other", text: "Other Site (Specify Url)" });
            for (const _key in response1.value) {
              if (this.context.pageContext.web.absoluteUrl != response1.value[_key]["Url"]) {
                sites.push({
                  key: response1.value[_key]["Url"],
                  text: response1.value[_key]["Title"],
                });
              }
            }
            this._siteOptions = sites;

            if (this.properties.siteUrl && this.properties.siteUrl !== 'other') {
              void this._getListTitles(this.properties.siteUrl).then((response2) => {
                this._listOptions = response2.value.map((list: any) => {
                  return {
                    key: list.Title,
                    text: list.Title,
                  };
                });
                this._listsDisabled = false;
                this.context.propertyPane.refresh();
                this.context.statusRenderer.clearLoadingIndicator(this.domElement);
                this.render();
              });
            } else {
              this.context.propertyPane.refresh();
              this.context.statusRenderer.clearLoadingIndicator(this.domElement);
              this.render();
            }
          });
      });
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'siteUrl' && newValue) {
      this._listsDisabled = true;
      this.properties.listName = "";
      this.context.propertyPane.refresh();

      let siteUrl = newValue;
      if (newValue === 'other') {
        // Valid 'other' URL will be handled by siteUrl field itself if we separate it, 
        // but here we might just treat 'other' as a selection that shows a text box.
        // For simplicity based on reference, if site is selected, load lists.
        this.context.propertyPane.refresh();
      }

      void this._getListTitles(siteUrl).then((response) => {
        this._listOptions = response.value.map((list: any) => {
          return {
            key: list.Title,
            text: list.Title,
          };
        });
        this._listsDisabled = false;
        this.context.propertyPane.refresh();
        this.render();
      });
    }
    else {
      this.render();
    }
  }

  private _getSiteRootWeb(): Promise<any> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl + `/_api/Site/RootWeb?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getSites(rootWebUrl: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        rootWebUrl + `/_api/web/webs?$select=Title,Url`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getListTitles(site: string): Promise<any> {
    return this.context.spHttpClient
      .get(
        site + `/_api/web/lists?$filter=Hidden eq false and BaseType eq 0`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
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
                }),
                PropertyPaneDropdown('siteUrl', {
                  label: "Site",
                  options: this._siteOptions
                }),
                PropertyPaneDropdown('listName', {
                  label: "List Name",
                  options: this._listOptions,
                  disabled: this._listsDisabled
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
