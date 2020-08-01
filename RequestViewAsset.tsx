import * as React from "react";
import styles from "./RequestNewAsset.module.scss";
import { IRequestNewAssetProps } from "./IRequestNewAssetProps";
import { IRequestNewAssetState } from "./IRequestNewAssetState";
import { sp, Web, List } from "@pnp/sp/presets/all";
import ReactTable from "react-table-v6";
import "react-table-v6/react-table.css";
import { escape } from "@microsoft/sp-lodash-subset";
import tablestyles from "./ReactTable.scss";
import {
  PrimaryButton,
  Icon,
  ActionButton,
  DefaultButton,
} from "office-ui-fabric-react";
import { initializeIcons } from "@uifabric/icons";
const columns = [
  {
    Header: (
      <PrimaryButton>
        <b>Edit Item</b>
      </PrimaryButton>
    ),
    Cell: (row) => (
      <div
        style={{
          width: "100%",
          height: "100%",
          backgroundColor: "#a36cf1"
            
        }}>
          {row.value}
      </div>
    ),
    accessor: "Edit",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        <b>Request ID</b>
      </PrimaryButton>
    ),
    accessor: "RequestID",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        <b>Request Status</b>
      </PrimaryButton>
    ),
    accessor: "RequestStatus",
    sortable: true,
    Cell: (row) => (
      <div
        style={{
          width: "100%",
          height: "100%",
          backgroundColor:
            row.value == "Approved"
              ? "Orange"
              : row.value == "Rejected"
              ? "red"
              : row.value == "Delivered"
              ? "forestgreen"
              : row.value == "Pending Approval"
              ? "Blue"
              : "gold",
          borderRadius: "2px",
          color: "white",
          textAlign: "center",
        }}
      >
        {row.value}
      </div>
    ),
  },
  {
    Header: (
      <PrimaryButton>
        {" "}
        <b>Approver Name</b>
      </PrimaryButton>
    ),
    accessor: "ApproverName",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        <b>Asset Type</b>
      </PrimaryButton>
    ),
    accessor: "AssetType",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        {" "}
        <b>Department</b>
      </PrimaryButton>
    ),
    accessor: "Department",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        <b>Purpose of Usage</b>
      </PrimaryButton>
    ),
    accessor: "PurposeofUsage",
    sortable: true,
  },
  {
    Header: (
      <PrimaryButton>
        <b>Reason for Rejection</b>
      </PrimaryButton>
    ),
    accessor: "ReasonforRejection",
    sortable: true,
  },

  {
    Header: (
      <PrimaryButton>
        <b>Requested On</b>
      </PrimaryButton>
    ),
    accessor: "RequestedOn",
    sortable: true,
  },
];
export default class RequestNewAsset extends React.Component<
  IRequestNewAssetProps,
  IRequestNewAssetState
> {
  constructor(props: IRequestNewAssetProps, state: IRequestNewAssetState) {
    super(props);
    initializeIcons();
    this.state = {
      RequestId: null,
      AssetType: "",
      PurposeofUsage: "",
      ApproverName: null,
      RequestStatus: "",
      RequestedOn: null,
      Department: "",
      Data: [],
    };
  }
  public componentDidMount() {
    debugger;
    initializeIcons();
    this.getLists();
  }
  public render(): React.ReactElement<IRequestNewAssetProps> {
    let { Data } = this.state;
    return (
      <div>
        <ReactTable
          data={Data}
          columns={columns}
          defaultPageSize={10}
          minRows={0}
          className="-stripedehighlight"
          SortIcon={tablestyles}
        />
      </div>
    );
  }
  private async GetLists() {
    let web = Web("https://chhowdam.sharepoint.com/sites/demo1/");
    return web.lists
      .getByTitle("IT Asset Requests Tracker")
      .items.select(
        "ID",
        "Request_x0020_ID",
        "Approver_x0020_Name/Title",
        "Asset_x0020_Type",
        "Department1",
        "Purpose_x0020_of_x0020_Usage",
        "Reason_x0020_for_x0020_Rejection",
        "Request_x0020_Status",
        "Requested_x0020_On"
      )
      .expand("Approver_x0020_Name")
      .getAll()
      .then((data) => {
        return data;
      });
  }


  
  private getLists() {
    let Values: any = [];
    debugger;
    this.GetLists().then((response) => {
     
      if (response.length > 0) {
        for (let i = 0; i < response.length; i++) {
          let href = `https://chhowdam.sharepoint.com/sites/demo1/SitePages/IT-Asset-Request.aspx?itemId=${response[i].ID}&FormMode=Edit`;
          //let edit=`<a href=${url}>Edit</a>`
          let date=new Date(response[i].Requested_x0020_On);
          Values.push({
            RequestID:
              response[i].Request_x0020_ID != null
                ? response[i].Request_x0020_ID
                : "",
            AssetType:
              response[i].Asset_x0020_Type != ""
                ? response[i].Asset_x0020_Type
                : "",
            PurposeofUsage:
              response[i].Purpose_x0020_of_x0020_Usage != ""
                ? response[i].Purpose_x0020_of_x0020_Usage
                : "",
            ApproverName:
              response[i].Approver_x0020_Name != null
                ? response[i].Approver_x0020_Name.Title
                : "",
            RequestStatus:
              response[i].Request_x0020_Status != ""
                ? response[i].Request_x0020_Status
                : "",
            RequestedOn:
              response[i].Requested_x0020_On != ""
                ? date.toLocaleDateString()
                : "",
            ReasonforRejection:
              response[i].Reason_x0020_for_x0020_Rejection != ""
                ? response[i].Reason_x0020_for_x0020_Rejection
                : "",
            Department:
              response[i].Department1 != "" ? response[i].Department1 : "",
            Edit: (
              <ActionButton
                data-automation-id="submit"
                iconProps={{ iconName: "Edit" }}
              >
                <a href={href} color={"WHITE"}>
                  Edit
                </a>
              </ActionButton>
            ),
          });
        }
        this.setState({
          Data: Values,
        });
      }
    });
  }
}
