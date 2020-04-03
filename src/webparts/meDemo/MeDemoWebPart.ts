import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MeDemoWebPartStrings';
import MeDemo from './components/MeDemo';
import { IMeDemoProps } from './components/IMeDemoProps';

export interface IMeDemoWebPartProps {
  description: string;
}

export default class MeDemoWebPart extends BaseClientSideWebPart <IMeDemoWebPartProps> {

  public render(): void {
    let isMicrosoftTeams: boolean = false;
    if (this.context.sdks.microsoftTeams) {
      isMicrosoftTeams = true;
    }
    const element: React.ReactElement<IMeDemoProps> = React.createElement(
      MeDemo,
      {
        isMicrosoftTeams: isMicrosoftTeams,
        msGraphClientFactory: this.context.msGraphClientFactory
      }
    );

    ReactDom.render(element, this.domElement);
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
