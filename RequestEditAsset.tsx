import * as React from "react";
import styles from "./RequestNewAsset.module.scss";
import { IRequestNewAssetProps } from "./IRequestNewAssetProps";
import { IRequestNewAssetState } from "./IRequestNewAssetState";
import {
  TextField,
  Label,
  PrimaryButton,
  ChoiceGroup,
  DatePicker,
} from "office-ui-fabric-react/lib";

import { escape } from "@microsoft/sp-lodash-subset";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp, Web, List } from "@pnp/sp/presets/all";
import { UrlQueryParameterCollection } from "@microsoft/sp-core-library";
export default class RequestNewAsset extends React.Component<
  IRequestNewAssetProps,
  IRequestNewAssetState
> {
  constructor(props: IRequestNewAssetProps, state: IRequestNewAssetState) {
    super(props);
    this.state = {
      RequestId: "",
      AssetType: "",
      PurposeofUsage: "",
      ApproverName: null,
      RequestStatus: "",
      RequestedOn: null,
      Department: "",
      ReasonforRejection: "",
      currentuserEmail:""
    };
  }
  public componentDidMount() {
    

    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let itemIDParm: string = queryParms.getValue("ItemId");
    this.checkUser(itemIDParm);
   
  }
 private checkUser(itemIDParm) {
   this.checkUserInID(itemIDParm).then((response)=>{
     let ApproverName=response.Approver_x0020_Name.UserName;
     if(ApproverName==this.props.currentUser["userName"]){
      this.getLists(itemIDParm);
     }
     else{
       alert("You dont have access to Edit the form.");
     }
   })
 }
 private checkUserInID(itemIDParm) {

   let web = Web("https://chhowdam.sharepoint.com/sites/demo1/");
   return web.lists
      .getByTitle("IT Asset Requests Tracker")
      .items.getById(itemIDParm)
      .select(
        "ID",
        "Request_x0020_ID",
        "Approver_x0020_Name/Id",
        "Approver_x0020_Name/Title",
        "Approver_x0020_Name/UserName",
        "Asset_x0020_Type",
        "Department1",
        "Purpose_x0020_of_x0020_Usage",
        "Reason_x0020_for_x0020_Rejection",
        "Request_x0020_Status",
        "Requested_x0020_On"
      )
      .expand("Approver_x0020_Name")
      .get()
      .then((data) => {
        return data;
      });
 }
  private async GetLists(itemIDParm) {
    let web = Web("https://chhowdam.sharepoint.com/sites/demo1/");
    return web.lists
      .getByTitle("IT Asset Requests Tracker")
      .items.getById(itemIDParm)
      .select(
        "ID",
        "Request_x0020_ID",
        "Approver_x0020_Name/Id",
        "Approver_x0020_Name/Title",
        "Approver_x0020_Name/UserName",
        "Asset_x0020_Type",
        "Department1",
        "Purpose_x0020_of_x0020_Usage",
        "Reason_x0020_for_x0020_Rejection",
        "Request_x0020_Status",
        "Requested_x0020_On"
      )
      .expand("Approver_x0020_Name")
      .get()
      .then((data) => {
        return data;
      });
  }
  public render(): React.ReactElement<IRequestNewAssetProps> {
    return (
      <>
        <div className={styles.requestNewAsset}>
          <div className={styles.container}>
            <div className={styles.grid}>
              <div className={styles.row}>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Request ID
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
                  >
                    <TextField
                      type="Number"
                   // disabled={true}
                      errorMessage={this.state.RequestId===null?"Its a mandatory field":""}
                      value={this.state.RequestId}
                      onChange={(
                        ev: React.FormEvent<HTMLInputElement>,
                        newValue?: string
                      ) => {
                        this.setState({
                          RequestId: newValue,
                        });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Asset Type
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-8 ms-lg8`}
                  >
                    <ChoiceGroup
                      styles={{ flexContainer: { display: "flex" } }}
                     
                    //  disabled
                      selectedKey={this.state.AssetType}
                      options={[
                        { key: "CTS Laptop", text: "CTS Laptop\u00A0\u00A0" },
                        { key: "Desktop", text: "Desktop\u00A0\u00A0" },
                        {
                          key: "Client Laptop",
                          text: "Client Laptop\u00A0\u00A0",
                        },
                        { key: "Others", text: "Others" },
                      ]}
                      onChange={(
                        ev: React.FormEvent<HTMLInputElement>,
                        choice: any
                      ) => {
                        this.setState({ AssetType: choice.key });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Purpose of Usage
                    </Label>
                  </div>

                  <div
                    className={`${styles.choiceColumn} ms-sm-12 ms-md-12 ms-lg8`}
                  >
                    <TextField
                      multiline
                      autoAdjustHeight
                      value={this.state.PurposeofUsage}
                 //    disabled={true}

                      onChange={(
                        ev: React.FormEvent<HTMLInputElement>,
                        newValue?: string
                      ) => {
                        this.setState({
                          PurposeofUsage: newValue,
                        });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Approver Name
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
                  >
                    <PeoplePicker
                      context={this.props.context}
                      groupName={""}
                      showtooltip={true}
                    
                   //  disabled={true}

                      ensureUser={true}
                 selectedItems={this.__people.bind(this)}
                      showHiddenInUI={false}
                      principalTypes={[PrincipalType.User]}
                      resolveDelay={1000}
                      personSelectionLimit={1}
                      defaultSelectedUsers={[this.state.currentuserEmail]}

                    />
                  </div>
                </div>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Request Status
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-8 ms-lg8 `}
                  >
                    <ChoiceGroup
                      required={true}
                      selectedKey={this.state.RequestStatus}
                      styles={{ flexContainer: { display: "flex" } }}
                      options={[
                        { key: "Approved", text: "Approved\u00A0\u00A0" },
                        { key: "Rejected", text: "Rejected\u00A0\u00A0" },
                        { key: "Delivered", text: "Delivered\u00A0\u00A0" },
                        {
                          key: "Pending Approval",
                          text: "Pending Approval\u00A0\u00A0",
                        },
                        {
                          key: "Scheduled For Delivery",
                          text: " Scheduled For Delivery",
                        },
                      ]}
                      onChange={(
                        ev: React.FormEvent<HTMLInputElement>,
                        choice: any
                      ) => {
                        this.setState({ RequestStatus: choice.key });
                      }}
                    />
                  </div>
                </div>
                {this.state.RequestStatus=="Rejected"?
              <div>
              <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                <Label
                  style={{
                    textAlign: "left",
                    paddingLeft: "0px",
                    fontSize: 13,
                    fontWeight: "bold",
                  }}
                >
                 Reason for Rejection
                </Label>
              </div>

              <div
                className={`${styles.choiceColumn} ms-sm-12 ms-md-12 ms-lg8`}
              >
                <TextField
                  multiline
                  autoAdjustHeight
                  value={this.state.ReasonforRejection}
                // disabled={true}

                  onChange={(
                    ev: React.FormEvent<HTMLInputElement>,
                    newValue?: string
                  ) => {
                    this.setState({
                      ReasonforRejection: newValue,
                    });
                  }}
                />
              </div>
            </div>:""
              }
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Requested On
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
                  >
                    <DatePicker
                      
                     
                    // disabled
                      placeholder="Select a date..."
                      onSelectDate={this._onSelectDate}
                      value={this.state.RequestedOn}
                      formatDate={this._onFormatDate}
                      isMonthPickerVisible={false}
                    />
                  </div>
                </div>
                <div>
                  <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                    <Label
                      style={{
                        textAlign: "left",
                        paddingLeft: "0px",
                        fontSize: 13,
                        fontWeight: "bold",
                      }}
                    >
                      Department
                    </Label>
                  </div>

                  <div
                    className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
                  >
                    <ChoiceGroup
                             
                      selectedKey={this.state.Department}
                      styles={{ flexContainer: { display: "flex" } }}
                      options={[
                        { key: "IT", text: "IT\u00A0\u00A0" },
                        { key: "BPO", text: "BPO\u00A0\u00A0" },
                        { key: "ITES", text: "ITES\u00A0\u00A0" },
                        { key: "INFRA", text: " INFRA" },
                      ]}
                      onChange={(
                        ev: React.FormEvent<HTMLInputElement>,
                        choice: any
                      ) => {
                        this.setState({ Department: choice.key });
                      }}
                    />
                  </div>
                </div>
                <div>
                  <PrimaryButton
                    text="Submit"
                      onClick={this.saveItem}
                    allowDisabledFocus
                    disabled={false}
                    checked={false}
                    styles={{
                      root: [
                        {
                          minWidth: 100,
                          marginRight: "20px",
                          height: "32px",
                          background: "#0078D4",
                          fontSize: "16px",
                          borderRadius: "2px",
                        },
                      ],
                      rootHovered: [
                        {
                          background: "#106EBE",
                          fontSize: 16,
                        },
                      ],
                    }}
                  />
                  <PrimaryButton
                    text="Close"
                    // onClick={this.Close}
                    allowDisabledFocus
                    disabled={false}
                    checked={false}
                    styles={{
                      root: [
                        {
                          minWidth: 100,
                          height: "32px",
                          background: "#0078D4",
                          fontSize: "16px",
                          borderRadius: "2px",
                        },
                      ],
                      rootHovered: [
                        {
                          background: "#106EBE",
                          fontSize: 16,
                        },
                      ],
                    }}
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      </>
    );
  }
  private _onSelectDate = (date: Date | null | undefined): void => {
    this.setState({ RequestedOn: date });
  }
  private _onFormatDate = (date: Date): string => {
    let newDate: string = null;
    let newMonth: string = null;
    if (date.getMonth() < 9) {
      newMonth = "0" + (date.getMonth() + 1);
    } else {
      newMonth = (date.getMonth() + 1).toString();
    }

    if (date.getDate() < 10) {
      newDate = "0" + date.getDate();
    } else {
      newDate = date.getDate().toString();
    }

    return newMonth + "/" + newDate + "/" + date.getFullYear();
  }
  private __people(items: any) {
    let currentComponent = this;
    let tempPplArr = 0;
    items.forEach((item) => {
      tempPplArr = item.id;
    });
    currentComponent.setState({
      ApproverName: tempPplArr,
    });
  }
  private getLists(itemIDParm) {
    let Values: any = [];
    debugger;
    this.GetLists(itemIDParm).then((response) => {
      let Approver = response.Approver_x0020_Name !== undefined ? response.Approver_x0020_Name.UserName : "";
      let ApproverId=response.Approver_x0020_Name !== undefined ? response.Approver_x0020_Name.Id : "";
      if (response) {
        this.setState({
          RequestId: response.Request_x0020_ID,
          AssetType: response.Asset_x0020_Type,
          PurposeofUsage: response.Purpose_x0020_of_x0020_Usage,
          ApproverName: ApproverId,
          RequestStatus: response.Request_x0020_Status,
          RequestedOn: this.fnParseDateFromString(
            response.Requested_x0020_On
          ),
          ReasonforRejection: response.Reason_x0020_for_x0020_Rejection,
          Department: response.Department1,
          currentuserEmail: Approver,
        });
      }
    });
  }
  private fnParseDateFromString = (value: string): Date => {
    const date = new Date();

    if (value !== null) {
      const values = (value || "").trim().split("-");
      const day =
        values.length > 0
          ? Math.max(1, Math.min(31, parseInt(values[2], 10)))
          : date.getDate();
      const month =
        values.length > 1
          ? Math.max(1, Math.min(12, parseInt(values[1], 10))) - 1
          : date.getMonth();
      let year =
        values.length > 2 ? parseInt(values[0], 10) : date.getFullYear();
      if (year < 100) {
        year += date.getFullYear() - (date.getFullYear() % 100);
      }
      return new Date(year, month, day);
    }
  }
  private saveItem = () => {
    let queryParms = new UrlQueryParameterCollection(window.location.href);
    let itemIDParm: any = queryParms.getValue("ItemId");
    const web=Web("https://chhowdam.sharepoint.com/sites/demo1");

    debugger;
    const item = web.lists
      .getByTitle("IT Asset Requests Tracker")
      .items.getById(itemIDParm).update({
       Request_x0020_ID:this.state.RequestId,
  Asset_x0020_Type:this.state.AssetType,
       Purpose_x0020_of_x0020_Usage:this.state.PurposeofUsage,
     Approver_x0020_NameId:this.state.ApproverName,
       Request_x0020_Status:this.state.RequestStatus,
    Requested_x0020_On:this.state.RequestedOn,
        Reason_x0020_for_x0020_Rejection:this.state.RequestStatus==="Rejected"?this.state.ReasonforRejection:"",
       Department1:this.state.Department,
   
      })
      .then((response) => {
        if(response){
          alert("Item is Updated Successfully!");
          window.location.href =
         "https://chhowdam.sharepoint.com/sites/demo1/SitePages/IT-Asset-Request.aspx?FormMode=View";
        }
      });
  }
}
