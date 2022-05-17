import { Label } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as React from "react";
import siteconfig from "../config/siteconfig.json";
import styles from "../scss/CMPApproveChampion.module.scss";
import * as LocaleStrings from 'ClbHomeWebPartStrings';



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
  list: ISPLists;
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
  }

  public componentDidMount(): void {
    this.setState({
      list: this.props.list
    });
  }

  private updateItem = (statusText: string, ID: number) => {
    let ButtonText = statusText;
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
        //filter updated item from state
        let filteredItems = this.state.list.value.filter((i: ISPList) => i.ID !== ID);
        if (response.status === 201) {
          this.setState({
            UserDetails: [],
            isAddChampion: false,
            list: { value: filteredItems }
          });
          alert("Champion" + status);
        } else {
          if (status === 'Approved') {
            this.setState({
              approveMessage: LocaleStrings.ChampionApprovedMessage,
              list: { value: filteredItems }
            });
          }
          if (status === 'Rejected') {
            this.setState({
              rejectMessage: LocaleStrings.ChampionRejectedMessage,
              list: { value: filteredItems }
            });
          }
        }
      });
  }

  public render() {
    return (
      <div className="container">
        <div className={styles.approveChampionPath}>
          <img src={require("../assets/CMPImages/BackIcon.png")}
            className={styles.backImg}
            alt={LocaleStrings.BackButton}
          />
          <span
            className={styles.backLabel}
            onClick={() => { this.props.onClickAddmember(this.state.list); }}
            title={LocaleStrings.CMPBreadcrumbLabel}
          >
            {LocaleStrings.CMPBreadcrumbLabel}
          </span>
          <span className={styles.border}></span>
          <span className={styles.approveChampionLabel}>{LocaleStrings.ManageApprovalsPageTitle}</span>
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
        <div className={styles.listHeading}>{LocaleStrings.ChampionsListPageTitle}</div>
        <table className="table table-bodered">
          <thead className={styles.listHeader}>
            <th>{LocaleStrings.PeopleNameGridHeader}</th>
            <th>{LocaleStrings.RegionGridHeader}</th>
            <th>{LocaleStrings.CountryGridHeader}</th>
            <th>{LocaleStrings.FocusAreaGridHeader}</th>
            <th>{LocaleStrings.GroupGridHeader}</th>
            {!this.props.isEmp && <th>{LocaleStrings.StatusGridHeader}</th>}
            <th>{LocaleStrings.ActionGridHeader}</th>
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
                          onClick={e => this.updateItem("Reject", item.ID)}
                          title={LocaleStrings.RejectButton}
                        >
                          <Icon iconName="ErrorBadge" className={`${classes.rejectIcon}`} />
                          <span className={styles.rejectBtnLabel}>{LocaleStrings.RejectButton}</span>
                        </button>
                        <button
                          className={`btn ${styles.approveBtn}`}
                          onClick={e => this.updateItem("Approve", item.ID)}
                          title={LocaleStrings.ApproveButton}
                        >
                          <Icon iconName="Completed" className={`${classes.approveIcon}`} />
                          <span className={styles.approveBtnLabel}>{LocaleStrings.ApproveButton}</span>
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
              <span className={styles.noRecordsLabels}>{LocaleStrings.NoChampionsMessage}</span>
            </div>
          )
        }

      </div>
    );
  }
}

export default ApproveChampion;
