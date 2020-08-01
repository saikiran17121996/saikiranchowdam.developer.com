import * as React from 'react';
import styles from './RequestNewAsset.module.scss';
import { IRequestNewAssetProps } from './IRequestNewAssetProps';
import { IRequestNewAssetState } from './IRequestNewAssetState';
import { TextField,Label ,PrimaryButton,ChoiceGroup,DatePicker} from 'office-ui-fabric-react/lib';
import { sp, Web, List } from "@pnp/sp/presets/all";
import { escape } from '@microsoft/sp-lodash-subset';
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
export default class RequestNewAsset extends React.Component<IRequestNewAssetProps,IRequestNewAssetState> {
constructor(props:IRequestNewAssetProps,state:IRequestNewAssetState){
  super(props);
  this.state={
    RequestId:"",
    AssetType:"",
    PurposeofUsage:"",
    ApproverName:null ,
    RequestStatus:"" ,
    RequestedOn:null ,
    Department:""
    };
  }
 public componentDidMount(){

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
                    fontWeight:"bold"
                    
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
                 required
                 value={this.state.RequestId}
                  onChange={(event:any, selectedOption) => {
                    this.setState({
                      RequestId:event.target.value
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
                    fontWeight:"bold"
                    
                  }}
                >
                Asset Type 
                </Label>
              </div>

              <div
                className={`${styles.mediumColumn} ms-sm6 ms-md-8 ms-lg8`}
              >
                <ChoiceGroup  
                 styles={{ flexContainer: { display: "flex"}}}
                 required
                 value={this.state.AssetType}
                options={[{ key: 'CTS Laptop', text: 'CTS Laptop\u00A0\u00A0' },
                         { key: 'Desktop', text: 'Desktop\u00A0\u00A0' },
                         { key: 'Client Laptop', text: 'Client Laptop\u00A0\u00A0',},
                         { key: 'Others', text: 'Others' }]} 
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
                    fontWeight:"bold"
                    
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
                required
                  onChange={(event:any, selectedOption) => {
                   this.setState({
                  PurposeofUsage:event.target.value
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
                    fontWeight:"bold"
                    
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
                          isRequired={true}
                          disabled={false}
                          ensureUser={true}
                          selectedItems={this.__people.bind(this)}
                          showHiddenInUI={false}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          personSelectionLimit={1}
                         // defaultSelectedUsers={[this.props.currentUser["emailId"]]}

                        />
              </div>
            </div>
            {/* <div>
              <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                <Label
                  style={{
                    textAlign: "left",
                    paddingLeft: "0px",
                    fontSize: 13,
                    fontWeight:"bold"
                    
                  }}
                >
               Request Status 
                </Label>
              </div>

              <div
                  className={`${styles.mediumColumn} ms-sm6 ms-md-8 ms-lg8`}
              >
                <ChoiceGroup  
                required={true}
                value={this.state.RequestStatus}
                 styles={{ flexContainer: { display: "flex"}}}
                options={[{ key: 'Approved', text: 'Approved\u00A0\u00A0' },
                         { key: 'Rejected', text: 'Rejected\u00A0\u00A0' },
                         { key: 'Delivered', text: 'Delivered\u00A0\u00A0'},
                         { key: 'Pending Approval', text: 'Pending Approval\u00A0\u00A0' },
                         { key: 'Scheduled For Delivery', text: ' Scheduled For Delivery' },
                        
                        ]}
                         onChange={(
                          ev: React.FormEvent<HTMLInputElement>,
                          choice: any
                        ) => {
                          this.setState({ RequestStatus: choice.key });
                        }}
                          />
              </div>
            </div> */}
            <div>
              <div className={`${styles.smallColumn} ms-sm6 ms-md6 ms-lg4`}>
                <Label
                  style={{
                    textAlign: "left",
                    paddingLeft: "0px",
                    fontSize: 13,
                    fontWeight:"bold"
                    
                  }}
                >
               Requested On
                </Label>
              </div>

              <div
                className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
              >
                <DatePicker
                          isRequired={true}
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
                    fontWeight:"bold"
                    
                  }}
                >
             Department
                </Label>
              </div>

              <div
                className={`${styles.mediumColumn} ms-sm6 ms-md-6 ms-lg8`}
              >
               <ChoiceGroup  
               required={true}
               value={this.state.Department}
                 styles={{ flexContainer: { display: "flex"}}}
                options={[{ key: 'IT', text: 'IT\u00A0\u00A0' },
                         { key: 'BPO', text: 'BPO\u00A0\u00A0' },
                         { key: 'ITES', text: 'ITES\u00A0\u00A0'},
                         { key: 'INFRA', text: ' INFRA' },
                        
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
    this.setState({ RequestedOn: date 
    });



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
  private __people(items: any) {
    let currentComponent = this;
    let tempPplArr = 0;
    items.forEach(item => {
      tempPplArr = item.id;
      });
      
      currentComponent.setState({
      ApproverName : tempPplArr,
    
    });
    
    
    }
    private saveItem = () => {
     
      debugger;
     const web=Web("https://chhowdam.sharepoint.com/sites/demo1")
      const item = web.lists
        .getByTitle("IT Asset Requests Tracker")
        .items.add({
         Request_x0020_ID:this.state.RequestId,
    Asset_x0020_Type:this.state.AssetType,
         Purpose_x0020_of_x0020_Usage:this.state.PurposeofUsage,
       Approver_x0020_NameId:this.state.ApproverName,
         Request_x0020_Status:this.state.RequestStatus,
      Requested_x0020_On:this.state.RequestedOn,
          Reason_x0020_for_x0020_Rejection:this.state.ReasonforRejection,
         Department1:this.state.Department,
     
        })
        .then((response) => {
          if(response){
            alert("Item Created!");
            window.location.href =
            "https://chhowdam.sharepoint.com/sites/demo1/Lists/IT%20Asset%20Requests%20Tracker/AllItems.aspx";
          }
        });
    }
}
