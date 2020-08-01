import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IRequestNewAssetProps {
 
  context?:WebPartContext;
  currentUser?:object;
}
