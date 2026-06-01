import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DeskBookingWebPartStrings';
import DeskBooking from './components/DeskBooking';
import { IDeskBookingProps } from './components/IDeskBookingProps';

export interface IDeskBookingWebPartProps {
  deskMasterListTitle: string;
  deskBookingListTitle: string;
}

export default class DeskBookingWebPart extends BaseClientSideWebPart<IDeskBookingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDeskBookingProps> = React.createElement(
      DeskBooking,
      {
        context: this.context,
        deskMasterListTitle: this.properties.deskMasterListTitle || 'Desk Master',
        deskBookingListTitle: this.properties.deskBookingListTitle || 'Desk Booking'
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
                PropertyPaneTextField('deskMasterListTitle', {
                  label: strings.DeskMasterListTitleLabel
                }),
                PropertyPaneTextField('deskBookingListTitle', {
                  label: strings.DeskBookingListTitleLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
