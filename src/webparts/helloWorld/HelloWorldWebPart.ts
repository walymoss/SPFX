import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string; // Si esto es para un PropertyPaneTextField
  walyprop1: string;   // Asumiendo que es un texto debido al PropertyPaneTextField
  walyprop2: boolean;  // Para un PropertyPaneCheckbox, se usa boolean
  walyprop3: string;   // Para un PropertyPaneDropdown, se usa string para almacenar la clave seleccionada
  walyprop4: boolean;  // Para un PropertyPaneToggle, se usa boolean
}


export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        <h3>Welcome to SharePoint Framework!</h3>
        <div>
        <span>${this.properties.walyprop1}</span>
        <span>${this.properties.walyprop2}</span>
        <span>${this.properties.walyprop1} + ${this.properties.walyprop2}</span>
        </div>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }



  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
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

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Configuración del WebPart"
          },
          groups: [
            {
              groupName: "Configuraciones Básicas",
              groupFields: [
                PropertyPaneTextField('walyprop1', {
                  label: "Texto Simple"
                }),
                PropertyPaneCheckbox('walyprop2', {
                  text: "Habilitar Opción"
                }),
                PropertyPaneDropdown('walyprop3', {
                  label: "Elija una Opción",
                  options: [
                    { key: '1', text: 'Opción 1' },
                    { key: '2', text: 'Opción 2' }
                  ]
                }),
                PropertyPaneToggle('walyprop4', {
                  label: "Activar Función",
                  onText: "Sí",
                  offText: "No"
                })
              ]
            }
          ]
        }
      ]
    };
  }
  
}
