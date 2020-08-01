import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { sp, Web, List } from "@pnp/sp/presets/all";

import * as strings from 'RequestNewAssetWebPartStrings';
import RequestNewAsset from './components/RequestNewAsset';
import { IRequestNewAssetProps } from './components/IRequestNewAssetProps';

export interface IRequestNewAssetWebPartProps {
  description: string;
 
}

export default class RequestNewAssetWebPart extends BaseClientSideWebPart <IRequestNewAssetWebPartProps> {

  public render(): void {
    let web = Web(this.context.pageContext.site.absoluteUrl);
    console.log("Log Test");

    web.currentUser.get().then(res => {
      console.log("Log Web::-");
    const element: React.ReactElement<IRequestNewAssetProps> = React.createElement(
      RequestNewAsset,
      {
       
        context:this.context,
        currentUser: {
          userName: res.Title,
          emailId: res.UserPrincipalName.toLowerCase(),
          userId: res.Id
        }
      }
    );

    ReactDom.render(element, this.domElement);
  });
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
