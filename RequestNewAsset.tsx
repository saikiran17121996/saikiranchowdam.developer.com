import * as React from "react";
import styles from "./RequestNewAsset.module.scss";
import { IRequestNewAssetProps } from "./IRequestNewAssetProps";
import { escape } from "@microsoft/sp-lodash-subset";
import RequestViewAsset from "./RequestViewAsset";
import RequestAsset from "./RequestAsset";
import RequestEditAsset from "./RequestEditAsset";
//Importing SPComponent Loader for dynamically loading files from SharePoint libraries
import { SPComponentLoader } from "@microsoft/sp-loader";
//Importing UrlQueryParameterCollection from SharePoint libraries
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";
export default class RequestNewAsset extends React.Component<
  IRequestNewAssetProps,
  {}
> {
  public render(): React.ReactElement<IRequestNewAssetProps> {
    debugger;
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let itemIDParm: string = queryParms.getValue("ItemId");
    let Mode: string = queryParms.getValue("FormMode");
    let itemId = itemIDParm === undefined || "" ? "" : itemIDParm;
    if (Mode === "New") {
      return (
        <div>
          <RequestAsset
            context={this.props.context}
            currentUser={this.props.currentUser}
          />
        </div>
      );
    } else if (itemId !== "" && Mode === "Edit") {
      return (
        <div>
          <RequestEditAsset
            context={this.props.context}
            currentUser={this.props.currentUser}
          />
        </div>
      );
    } else {
      return (
        <div>
          <RequestViewAsset
            context={this.props.context}
            currentUser={this.props.currentUser}
          />
        </div>
      );
    }
  }
}
