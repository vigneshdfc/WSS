import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'WorksheetWebPartStrings';
import Worksheet from './components/Worksheet'; // React component
import { IWorksheetProps } from './components/IWorksheetProps'; // Props for the component

// 👇 This is ONLY for web part properties (property pane)
export interface IWorksheetWebPartProps {
  description: string;
}

export default class WorksheetWebPart extends BaseClientSideWebPart<IWorksheetWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IWorksheetProps> = React.createElement(
      Worksheet, // React component
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      const style = document.createElement("style");
      style.innerHTML = `
      @media screen and (min-width: 1024px) {
      .c_b_cb6f7c2e:not(.g_b_cb6f7c2e) .j_b_cb6f7c2e {
        display: block !important;
        max-width: 1600px !important;
        }
      }
     
      .CanvasZone {
        padding: 0  !important;
        max-width: 1600px !important;
        margin: 0 auto !important;
      }
     
      .CanvasSection {
        padding: 0 !important;
      }
     
      .ControlZone .ControlZone--clean .a_a_50a7110f {
        margin: 0 !important;
        padding: 0 !important;
      }
       
      #workbenchPageContent {
        max-width: none !important;
        width: 100% !important;
      }
     
      #spLeftNav {
        display: none;
      }
 
      .r_NLtZH_y298L:not(.f_bHim3_y298L) .s_wDEw-_y298L {
        max-width: none !important;
      }
 
      #spSiteHeader{
        display: none;
      }
       
      #sp-appBar,
      #spCommandBar,
      #SuiteNavWrapper,
      .commandBarButtonHeightAndColor[aria-label="Command bar"],
      .headerRow-50,
      .j_m_4ade22aa#CommentsWrapper {
        display: none !important;
      }
       
      .p_e_8474018e {
        padding: 0 !important;
      }
       
      .a_g_cb6f7c2e:not(.e_g_cb6f7c2e):not(.aa_g_cb6f7c2e) {
        margin: 0 !important;
        padding: 0 !important;
      }
 
      @media screen and (min-width: 1024px) {
        .a_a_cb6f7c2e:not(.e_a_cb6f7c2e) .h_a_cb6f7c2e {
          max-width: 100% !important;
        }
      }
        .m_b_d71df89c {
        overflow-y: hidden !important;
        }`
        document.head.appendChild(style);
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
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
