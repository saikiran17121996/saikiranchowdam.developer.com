export interface IRequestNewAssetState {
    description?: string;
    RequestId?:string;
    AssetType?:string;
    PurposeofUsage?:string;
    ApproverName?:any ;
    RequestStatus?:string ;
    RequestedOn?:Date ;
    Department?:string;
    Data?:any[];
    ReasonforRejection?:string;
    currentuserEmail?:string;
  }
  