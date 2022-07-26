import { Label } from "@fluentui/react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as React from "react";
import siteconfig from "../config/siteconfig.json";
import styles from "../scss/CMPChampionsList.module.scss";
import * as LocaleStrings from 'ClbHomeWebPartStrings';


export interface IClbChampionsListProps {
  context?: WebPartContext;
  onClickAddmember: Function;
  isEmp: boolean;
  siteUrl: string;
  userAdded: boolean;
  userStatus: string;
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
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  successMessage: string;
  userDetails: Array<any>;
  selectedUsers: Array<any>;
  siteUrl: string;
  inclusionPath: string;
  siteName: string;

}
class ClbChampionsList extends React.Component<IClbChampionsListProps, IState> {
  constructor(props: IClbChampionsListProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });

    this.state = {
      list: { value: [] },
      isAddChampion: false,
      successMessage: "",
      userDetails: [],
      selectedUsers: [],
      siteUrl: this.props.siteUrl,
      siteName: siteconfig.sitename,
      inclusionPath: siteconfig.inclusionPath,
    };
    this._getListData();
  }

  //Get Details of all members from Member List 
  private _getListData(): Promise<ISPLists> {
    return this.props.context.spHttpClient
      .get("/" + this.state.inclusionPath + "/" + this.state.siteName +

        "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          let res = response.json();
          res.then((responseJSON: any) => {
            this._renderList(responseJSON.value);
          });
          return res;
        }
      });
  }

  private _renderList(items: ISPList[]): void {
    this.setState({ list: { value: items } });
  }

  public render() {
    return (
      <div className={`container ${styles.championListContainer}`}>
        <div className={styles.championListPath}>
          <img src={require("../assets/CMPImages/BackIcon.png")}
            className={styles.backImg}
            alt={LocaleStrings.BackButton}
          />
          <span
            className={styles.backLabel}
            onClick={() => { this.props.onClickAddmember(); }}
            title={LocaleStrings.CMPBreadcrumbLabel}
          >
            {LocaleStrings.CMPBreadcrumbLabel}
          </span>
          <span className={styles.border}></span>
          <span className={styles.championListLabel}>{LocaleStrings.ChampionsListPageTitle}</span>
        </div>
        {this.props.userAdded ?
          <Label className={styles.successMessage}>
            <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
            {this.props.userStatus === "Pending" ? LocaleStrings.UserNominatedMessage : LocaleStrings.UserAddedMessage}
          </Label> : null}
        <div className={`${styles.listHeading}`}>{LocaleStrings.ChampionsListPageTitle}</div>
        <div className={styles.championListTableArea}>
          <table className="table table-bodered">
            <thead className={styles.listHeader}>
              <th title={LocaleStrings.PeopleNameGridHeader}>{LocaleStrings.PeopleNameGridHeader}</th>
              <th title={LocaleStrings.RegionGridHeader}>{LocaleStrings.RegionGridHeader}</th>
              <th title={LocaleStrings.CountryGridHeader}>{LocaleStrings.CountryGridHeader}</th>
              <th title={LocaleStrings.FocusAreaGridHeader}>{LocaleStrings.FocusAreaGridHeader}</th>
              <th title={LocaleStrings.GroupGridHeader}>{LocaleStrings.GroupGridHeader}</th>
              {!this.props.isEmp && <th title={"Status"}>Status</th>}
            </thead>
            <tbody className={styles.listBody}>
              {this.state.list &&
                this.state.list.value &&
                this.state.list.value.length > 0 &&
                this.state.list.value.map((item: ISPList) => {
                  if (item.Status === "Approved") {//showing only approved list
                    return (
                      <tr>
                        <td title={`${item.FirstName ? item.FirstName + " " : ""}${item.LastName ? item.LastName : ""}`}>
                          {item.FirstName}
                          <span className="mr-1"></span>
                          {item.LastName}
                        </td>
                        <td title={item.Region ? item.Region : ""}>{item.Region}</td>
                        <td title={`${item.Country ? item.Country : ""}`}>{item.Country}</td>
                        <td title={`${item.FocusArea ? item.FocusArea : ""}`}>{`${item.FocusArea ? item.FocusArea : ""}`}</td>
                        <td title={`${item.Group ? item.Group : ""}`}>{item.Group}</td>
                        {!this.props.isEmp && <td title={`${item.Status ? item.Status : ""}`}>{item.Status}</td>}
                      </tr>
                    );
                  }
                })}
            </tbody>
          </table>
        </div>
      </div>
    );
  }
}

export default ClbChampionsList;
