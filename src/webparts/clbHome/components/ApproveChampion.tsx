import { Checkbox, Label, Spinner, SpinnerSize } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as React from "react";
import styles from "../scss/CMPApproveChampion.module.scss";
import commonServices from '../Common/CommonServices';
import * as stringsConstants from "../constants/strings";

let commonServiceManager: commonServices;

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
  approveMessage: string;
  rejectMessage: string;
  selectedIds: any;
  isAllSelected: boolean;
  showSpinner: boolean;
}
class ApproveChampion extends React.Component<IClbChampionsListProps, IState> {
  constructor(props: IClbChampionsListProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });
    this.state = {
      list: { value: [] },
      approveMessage: "",
      rejectMessage: "",
      selectedIds: [],
      isAllSelected: false,
      showSpinner: false
    };
    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
    this.setSelectedIds = this.setSelectedIds.bind(this);
    this.handleSelectAll = this.handleSelectAll.bind(this);
  }

  //Method will be called immediately after the component is mounted in DOM
  public componentDidMount(): void {
    //Get all the pending items from Member list
    this.getPendingItems();
  }

  //Method to retrieve pending items from Member list
  private async getPendingItems() {
    try {
      //Getting the pending items from Member List
      let filterQuery = "Status eq '" + stringsConstants.pendingStatus + "'";
      const pendingItems: any[] = await commonServiceManager.getItemsWithOnlyFilter(stringsConstants.MemberList, filterQuery);
     
      this.setState({
        list: { value: pendingItems }
      });
    }
    catch (error) {
      console.error("CMP_ApproveChampion_getPendingItems \n", JSON.stringify(error));
    }
  }

  //Method to update the status in the Member List
  private updatePendingItems = async (statusText: string, selectedIDs: any) => {
    try {
      this.setState({
        showSpinner: true,
        rejectMessage: "",
        approveMessage: ""
      });

      let updateMemberObject: any = {
        Status: statusText
      };

      //Update status for pending items in Member List
      let updateResponse = await commonServiceManager.updateMultipleItems(stringsConstants.MemberList, updateMemberObject, selectedIDs);

      //Refresh the items shown in the grid
      await this.getPendingItems();

      if (updateResponse) {
        //Updating state variables based on the approval action
        if (statusText === stringsConstants.approvedStatus) {
          this.setState({
            approveMessage: LocaleStrings.ChampionApprovedMessage,
            selectedIds: [],
            rejectMessage: "",
            isAllSelected: false,
            showSpinner: false
          });
        } else if (statusText === stringsConstants.rejectedStatus) {
          this.setState({
            rejectMessage: LocaleStrings.ChampionRejectedMessage,
            selectedIds: [],
            approveMessage: "",
            isAllSelected: false,
            showSpinner: false
          });
        }
      } else {
        //If any error occurs during batch update
        this.setState({
          rejectMessage: stringsConstants.CMPErrorMessage + "while approving/rejecting champion request(s).",
          selectedIds: [],
          isAllSelected: false,
          showSpinner: false
        });
      }
    }
    catch (error) {
      //Refresh the items shown in the grid
      await this.getPendingItems();

      this.setState({
        rejectMessage: stringsConstants.CMPErrorMessage + "while approving/rejecting champion request(s). Below are the details: \n" + JSON.stringify(error),
        selectedIds: [],
        isAllSelected: false,
        showSpinner: false
      });
      console.error("CMP_ApproveChampion_updatePendingItems \n", JSON.stringify(error));
    }
  }

  //Updating the state whenever the checkbox value is changed
  public setSelectedIds = (key: any, isChecked: boolean) => {
    if (isChecked) {
      this.setState({ selectedIds: [...this.state.selectedIds, key] });
      // Automatically check the "Select All" option when the last checkbox is checked
      if (this.state.selectedIds.length === this.state.list.value.length - 1) {
        this.setState({ isAllSelected: true });
      }
    }
    else {
      this.setState({
        isAllSelected: false,
        selectedIds: this.state.selectedIds.filter((tKey) => tKey !== key)
      });
    }
  }

  //Updating the state whenever the "Select All" checkbox value is changed
  public handleSelectAll = (isChecked: boolean) => {
    const tempArray: any = [];

    if (isChecked) {
      this.state.list.value.forEach((item: ISPList) => {
        tempArray.push(item.ID);
      });
      this.setState({
        isAllSelected: isChecked,
        selectedIds: tempArray
      });
    }
    else {
      this.setState({
        isAllSelected: isChecked,
        selectedIds: []
      });
    }
  }


  public render() {
    return (
      <div className={`container ${styles.approveChampionContainer}`}>
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
        <div className={styles.approveChampionTableArea}>
          <table className="table table-bodered">
            <thead className={styles.listHeader}>
              <th title={LocaleStrings.SelectAll}>
                <Checkbox
                  onChange={(eve: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked: boolean) => {
                    this.handleSelectAll(isChecked);
                  }}
                  checked={this.state.isAllSelected}
                  disabled={this.state.list.value.length == 0}
                />
              </th>
              <th title={LocaleStrings.PeopleNameGridHeader}>{LocaleStrings.PeopleNameGridHeader}</th>
              <th title={LocaleStrings.RegionGridHeader}>{LocaleStrings.RegionGridHeader}</th>
              <th title={LocaleStrings.CountryGridHeader}>{LocaleStrings.CountryGridHeader}</th>
              <th title={LocaleStrings.FocusAreaGridHeader}>{LocaleStrings.FocusAreaGridHeader}</th>
              <th title={LocaleStrings.GroupGridHeader}>{LocaleStrings.GroupGridHeader}</th>
              {!this.props.isEmp && <th>{LocaleStrings.StatusGridHeader}</th>}
            </thead>
            <tbody className={styles.listBody}>
              {this.state.list &&
                this.state.list.value &&
                this.state.list.value.length > 0 &&
                this.state.list.value.map((item: ISPList) => {
                  return (
                    <tr>
                      <td>
                        <Checkbox
                          value={item.ID}
                          onChange={(eve: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked: boolean) => {
                            this.setSelectedIds(item.ID, isChecked);
                          }}
                          checked={this.state.selectedIds.length > 0 ? this.state.selectedIds.includes(item.ID) : false}
                        />
                      </td>
                      <td title={`${item.FirstName ? item.FirstName + " " : ""}${item.LastName ? item.LastName : ""}`}>
                        {item.FirstName}
                        <span className="mr-1"></span>
                        {item.LastName}
                      </td>
                      <td title={item.Region ? item.Region : ""}>{item.Region}</td>
                      <td title={`${item.Country ? item.Country : ""}`}>{item.Country}</td>
                      <td title={`${item.FocusArea ? item.FocusArea : ""}`}>{`${item.FocusArea ? item.FocusArea : ""}`}</td>
                      <td title={`${item.Group ? item.Group : ""}`}>{item.Group}</td>
                      {!this.props.isEmp && <td>{item.Status}</td>}
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
        <div>
          {this.state.showSpinner &&
            <Spinner
              label={LocaleStrings.ProcessingSpinnerLabel}
              size={SpinnerSize.large}
            />
          }
        </div>
        {this.state.list.value.length > 0 &&
          <div className={styles.manageChampionBtnArea}>
            <button
              className={`btn ${styles.approveBtn}`}
              onClick={e => this.updatePendingItems(stringsConstants.approvedStatus, this.state.selectedIds)}
              title={LocaleStrings.ApproveButton}
              disabled={this.state.selectedIds.length === 0}
            >
              <Icon iconName="Completed" className={styles.approveBtnIcon} />
              <span className={styles.approveBtnLabel}>{LocaleStrings.ApproveButton}</span>
            </button>
            <button
              className={"btn " + styles.rejectBtn}
              onClick={e => this.updatePendingItems(stringsConstants.rejectedStatus, this.state.selectedIds)}
              title={LocaleStrings.RejectButton}
              disabled={this.state.selectedIds.length === 0}
            >
              <Icon iconName="ErrorBadge" className={styles.rejectBtnIcon} />
              <span className={styles.rejectBtnLabel}>{LocaleStrings.RejectButton}</span>
            </button>
          </div>
        }
        {this.state.list &&
          this.state.list.value &&
          this.state.list.value.length == 0 &&
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
