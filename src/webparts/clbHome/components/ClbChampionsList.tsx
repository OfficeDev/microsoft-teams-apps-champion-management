import { Label } from "@fluentui/react";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as React from "react";
import siteconfig from "../config/siteconfig.json";
import styles from "../scss/CMPChampionsList.module.scss";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { IConfigList } from "./ManageConfigSettings";
import * as stringConstants from "../constants/strings";
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";


export interface IClbChampionsListProps {
  context?: WebPartContext;
  onHomeCallBack: Function;
  siteUrl: string;
  userAdded: boolean;
  userStatus: string;
  configListData: Array<IConfigList>;
  memberListColumnsNames: Array<any>;
  appTitle: string;
  currentThemeName?: string
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
export interface IState {
  list: ISPLists;
  successMessage: string;
  siteUrl: string;
  inclusionPath: string;
  siteName: string;
  regionColumnName: string;
  countryColumnName: string;
  groupColumnName: string;
}
class ClbChampionsList extends React.Component<IClbChampionsListProps, IState> {
  constructor(props: IClbChampionsListProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any,
    });

    this.state = {
      list: { value: [] },
      successMessage: "",
      siteUrl: this.props.siteUrl,
      siteName: siteconfig.sitename,
      inclusionPath: siteconfig.inclusionPath,
      regionColumnName: "",
      countryColumnName: "",
      groupColumnName: ""
    };
    //Bind methods
    this.populateColumnNames = this.populateColumnNames.bind(this);
    this._getListData();
  }

  //Populate member list column display names 
  public componentDidMount(): void {
    this.populateColumnNames();
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

  //update component state with list data
  private _renderList(items: ISPList[]): void {
    this.setState({ list: { value: items } });
  }

  //Assign states with member list column names
  private populateColumnNames() {
    const enabledSettingsArray = this.props.configListData.filter((setting) => setting.Value === stringConstants.EnabledStatus);
    for (let setting of enabledSettingsArray) {
      const columnObject = this.props.memberListColumnsNames.find((column) => column.InternalName === setting.Title);
      if (columnObject.InternalName === stringConstants.RegionColumn) {
        this.setState({ regionColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringConstants.CountryColumn) {
        this.setState({ countryColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringConstants.GroupColumn) {
        this.setState({ groupColumnName: columnObject.Title });
      }
    }
  }

  public render() {
    const isDarkOrContrastTheme = this.props.currentThemeName === stringConstants.themeDarkMode || this.props.currentThemeName === stringConstants.themeContrastMode;
    return (
      <div className={`container ${styles.championListContainer}`}>
        <div className={`${styles.championListPath}${isDarkOrContrastTheme ? " " + styles.championListPathDarkContrast : ""}`}>
          <img src={require("../assets/CMPImages/BackIcon.png")}
            className={styles.backImg}
            alt={LocaleStrings.BackButton}
            aria-hidden="true"
          />
          <span
            className={styles.backLabel}
            onClick={() => { this.props.onHomeCallBack(); }}
            role="button"
            tabIndex={0}
            onKeyDown={(evt: any) => { if (evt.key === stringConstants.stringEnter) this.props.onHomeCallBack(); }}
            aria-label={this.props.appTitle}
          >
            <span title={this.props.appTitle}>
              {this.props.appTitle}
            </span>
          </span>
          <span className={styles.border}></span>
          <span className={styles.championListLabel}>{LocaleStrings.ChampionsListPageTitle}</span>
        </div>
        {this.props.userAdded ?
          <Label className={`${styles.successMessage}${isDarkOrContrastTheme ? " " + styles.successMessageDarkContrast : ""}`} aria-live="polite" role="alert">
            <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
            {this.props.userStatus === "Pending" ? LocaleStrings.UserNominatedMessage : LocaleStrings.UserAddedMessage}
          </Label> : null}
        <div className={`${styles.listHeading}${isDarkOrContrastTheme ? " " + styles.listHeadingDarkContrast : ""}`}>{LocaleStrings.ChampionsListPageTitle}</div>
        <div className={styles.championListTableArea}>
          <table className="table table-bodered">
            <thead className={styles.listHeader}>
              <th title={LocaleStrings.PeopleNameGridHeader}>{LocaleStrings.PeopleNameGridHeader}</th>
              {this.state.regionColumnName !== "" && <th title={this.state.regionColumnName}>{this.state.regionColumnName}</th>}
              {this.state.countryColumnName !== "" && <th title={this.state.countryColumnName}>{this.state.countryColumnName}</th>}
              {this.state.groupColumnName !== "" && <th title={this.state.groupColumnName}>{this.state.groupColumnName}</th>}
              <th title={LocaleStrings.FocusAreaGridHeader}>{LocaleStrings.FocusAreaGridHeader}</th>
            </thead>
            <tbody className={styles.listBody}>
              {this.state.list &&
                this.state.list.value &&
                this.state.list.value.length > 0 &&
                this.state.list.value.map((item: ISPList) => {
                  if (item.Status === "Approved") {//showing only approved list
                    return (
                      <tr>
                        <td>
                          <Person
                            personQuery={item.Title}
                            view={3}
                            personCardInteraction={1}
                            className="champion-person-card"
                          />
                        </td>
                        {this.state.regionColumnName !== "" && <td title={item.Region ? item.Region : ""}>{item.Region}</td>}
                        {this.state.countryColumnName !== "" && <td title={`${item.Country ? item.Country : ""}`}>{item.Country}</td>}
                        {this.state.groupColumnName !== "" && <td title={`${item.Group ? item.Group : ""}`}>{item.Group}</td>}
                        <td title={`${item.FocusArea ? item.FocusArea : ""}`}>{`${item.FocusArea ? item.FocusArea : ""}`}</td>
                      </tr>
                    );
                  }
                })}
            </tbody>
          </table>
        </div>
      </div >
    );
  }
}

export default ClbChampionsList;
