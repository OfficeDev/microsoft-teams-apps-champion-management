import { Label } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as React from "react";
import siteconfig from "../config/siteconfig.json";
import styles from "../scss/CMPApproveChampion.module.scss";



const classes = mergeStyleSets({
  rejectIcon: {
    marginRight: "10px",
    fontSize: "17px",
    fontWeight: "bolder",
    color: "#000003",
    opacity: 1
  },
  approveIcon: {
    marginRight: "10px",
    fontSize: "17px",
    fontWeight: "bolder",
    color: "#FFFFFF",
    opacity: 1
  }
});

export interface IClbChampionsListProps {
  context?: WebPartContext;
  onClickAddmember: Function;
  isEmp: boolean;
  siteUrl: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  FirstName: string;
  LastName: string;
  Country: String;
  Status: String;
  FocusArea: String;
  Group: String;
  Role: String;
  Region: string;
  Points: number;
  ID: number;
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  approveMessage: string;
  rejectMessage: string;
  UserDetails: Array<any>;
  selectedusers: Array<any>;
  siteUrl: string;
  memberrole: string;
}
class ApproveChampion extends React.Component<IClbChampionsListProps, IState> {
  constructor(props: IClbChampionsListProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });

    this.state = {
      list: { value: [] },
      isAddChampion: false,
      approveMessage: "",
      rejectMessage: "",
      UserDetails: [],
      selectedusers: [],
      siteUrl: this.props.siteUrl,
      memberrole: "",
    };
    this._getListData();
  }

  //Get the list of Members from member List
  private _getListData(): Promise<ISPLists> {
    return this.props.context.spHttpClient
      .get(
        "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          response.json().then((responseJSON: any) => {
            this._renderList(responseJSON.value);
          });
          return response.json();
        }
      });
  }

  private _renderList(items: ISPList[]): void {
    this.setState({ list: { value: items } });
  }

  private updateItem = (e, ID: number) => {
    let ButtonText = e.target.outerText;
    let status = "";
    let Id = ID;
    if (ButtonText === "Approve") {
      status = "Approved";
    }
    else {
      status = "Rejected";
    }
    const listDefinition: any = {
      Status: status,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listDefinition),
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
    };

    const url: string =
      "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + `/_api/web/lists/GetByTitle('Member List')/items(${Id})`;
    this.props.context.spHttpClient
      .post(
        url,
        SPHttpClient.configurations.v1,

        spHttpClientOptions
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 201) {
          this.setState({
            UserDetails: [],
            isAddChampion: false,
          });
          alert("Champion" + status);
        } else {
          if (status === 'Approved') {
            this.setState({
              approveMessage: `Your response has been ${status}.`
            });
          }
          if (status === 'Rejected') {
            this.setState({
              rejectMessage: `Your response has been ${status}.`
            });
          }
          this._getListData();
        }
      });
  }

  public render() {
    return (
      <div className="container">
        <div className={styles.approveChampionPath}>
          <img src={require("../assets/CMPImages/BackIcon.png")}
            className={styles.backImg}
          />
          <span
            className={styles.backLabel}
            onClick={() => { this.props.onClickAddmember(); }}
            title="Back"
          >
            Back
          </span>
          <span className={styles.border}></span>
          <span className={styles.approveChampionLabel}>Manage Approval</span>
        </div>
        {this.state.approveMessage &&
          <Label className={styles.approveMessage}>
            <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
            {this.state.approveMessage}
          </Label>
        }
        {this.state.rejectMessage &&
          <Label className={styles.rejectMessage}>
            {this.state.rejectMessage}
          </Label>
        }
        <div className={styles.listHeading}>Champion List</div>
        <table className="table table-bodered">
          <thead className={styles.listHeader}>
            <th>People Name</th>
            <th>Region</th>
            <th>Country</th>
            <th>FocusArea</th>
            <th>Group</th>
            {!this.props.isEmp && <th>Status</th>}
            <th>Action</th>
          </thead>
          <tbody className={styles.listBody}>
            {this.state.list &&
              this.state.list.value &&
              this.state.list.value.length > 0 &&
              this.state.list.value.map((item: ISPList) => {
                if (item.Status != "Approved" && item.Status != "Rejected") {//showing only approved list
                  return (
                    <tr>
                      <td>
                        {item.FirstName}
                        <span className="mr-1"></span>
                        {item.LastName}
                      </td>
                      <td>{item.Region}</td>
                      <td>{item.Country}</td>
                      <td>{item.FocusArea}</td>
                      <td>{item.Group}</td>
                      {!this.props.isEmp && <td>{item.Status}</td>}
                      <td>
                        <button
                          className={`btn ${styles.rejectBtn}`}
                          onClick={e => this.updateItem(e, item.ID)}
                          title="Reject"
                        >
                          <Icon iconName="ErrorBadge" className={`${classes.rejectIcon}`} />
                          <span className={styles.rejectBtnLabel}>Reject</span>
                        </button>
                        <button
                          className={`btn ${styles.approveBtn}`}
                          onClick={e => this.updateItem(e, item.ID)}
                          title="Approve"
                        >
                          <Icon iconName="Completed" className={`${classes.approveIcon}`} />
                          <span className={styles.approveBtnLabel}>Approve</span>
                        </button>
                      </td>
                    </tr>
                  );
                }
              })}
          </tbody>
        </table>
        {this.state.list &&
          this.state.list.value &&
          this.state.list.value.length > 0 &&
          this.state.list.value.filter(i => i.Status == "Pending").length == 0 &&
          (
            <div className={styles.noRecordsArea}>
              <img
                src={require('../assets/CMPImages/Norecordsicon.svg')}
                alt="norecordsicon"
                className={styles.noRecordsImg}
              />
              <span className={styles.noRecordsLabels}>NO CHAMPION REQUESTS AVAILABLE</span>
            </div>
          )
        }

      </div>
    );
  }
}

export default ApproveChampion;
