import { Checkbox, Label, SearchBox, Spinner, SpinnerSize } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as React from "react";
import BootstrapTable from 'react-bootstrap-table-next';
import paginationFactory from 'react-bootstrap-table2-paginator';
import ToolkitProvider, { ToolkitContextType } from 'react-bootstrap-table2-toolkit';
import commonServices from '../Common/CommonServices';
import * as stringConstants from "../constants/strings";
import styles from "../scss/ManageApprovals.module.scss";
import { IConfigList } from './ManageConfigSettings';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

let commonServiceManager: commonServices;

export interface IClbChampionsListProps {
  context: WebPartContext;
  siteUrl: string;
  setState: Function;
}
export interface ISPList {
  Title: string;
  FirstName: string;
  LastName: string;
  Country: string;
  Status: string;
  FocusArea: string;
  Group: string;
  Role: string;
  Region: string;
  Points: number;
  ID: number;
}
interface IState {
  championList: ISPList[];
  filteredChampionList: ISPList[];
  approveMessage: string;
  rejectMessage: string;
  selectedChampions: any;
  isAllSelected: boolean;
  showSpinner: boolean;
  configListSettings: Array<IConfigList>;
  memberListColumnNames: Array<any>;
  regionColumnName: string;
  countryColumnName: string;
  groupColumnName: string;
}
class ApproveChampion extends React.Component<IClbChampionsListProps, IState> {

  constructor(props: IClbChampionsListProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any,
    });
    this.state = {
      championList: [],
      filteredChampionList: [],
      approveMessage: "",
      rejectMessage: "",
      selectedChampions: [],
      isAllSelected: false,
      showSpinner: false,
      configListSettings: [],
      memberListColumnNames: [],
      regionColumnName: "",
      countryColumnName: "",
      groupColumnName: ""

    };
    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );

    //Bind Methods
    this.selectChampions = this.selectChampions.bind(this);
    this.getTableHeaderClass = this.getTableHeaderClass.bind(this);
    this.populateColumnNames = this.populateColumnNames.bind(this);
    this.getTableBodyClass = this.getTableBodyClass.bind(this);
  }

  //Method will be called immediately after the component is mounted in DOM
  //On component load show the pending members list
  public async componentDidMount() {
    //Get all the pending items from Member list
    await this.getPendingItems();

    //Get config list settings and memberlist display names
    await this.getConfigListSettings();
    await this.getMemberListColumnNames();
  }

  //Component did update life cycle method
  public componentDidUpdate(prevProps: Readonly<IClbChampionsListProps>, prevState: Readonly<IState>, snapshot?: any): void {

    //update column states with member list column display names 
    if (prevState.configListSettings !== this.state.configListSettings ||
      prevState.memberListColumnNames !== this.state.memberListColumnNames) {
      if (this.state.configListSettings.length > 0 && this.state.memberListColumnNames.length > 0)
        this.populateColumnNames();
    }
    //updating state of the parent component 'ManageApprovals" to show/hide the notification icon
    if (prevState.championList !== this.state.championList) {
      if (this.state.championList.length === 0) {
        this.props.setState({
          isPendingChampionApproval: false
        });
      }
      else {
        this.props.setState({
          isPendingChampionApproval: true
        });
      }
    }

  }

  //Get settings from config list
  private async getConfigListSettings() {
    try {
      const configListData: IConfigList[] = await commonServiceManager.getMemberListColumnConfigSettings();

      if (configListData.length === 3) {
        this.setState({ configListSettings: configListData });
      }
      else {
        this.setState({
          rejectMessage:
            stringConstants.CMPErrorMessage +
            ` while loading the page. There could be a problem with the ${stringConstants.ConfigList} data.`
        });
      }
    }
    catch (error) {
      console.error("CMP_ApproveChampion_getConfigListSettings \n", error);
      this.setState({
        rejectMessage:
          stringConstants.CMPErrorMessage +
          `while retrieving the ${stringConstants.ConfigList} settings. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //Get memberlist column names from member list
  private async getMemberListColumnNames() {
    try {
      const columnsDisplayNames: any[] = await commonServiceManager.getMemberListColumnDisplayNames();
      if (columnsDisplayNames.length > 0) {
        this.setState({ memberListColumnNames: columnsDisplayNames });
      }
    }
    catch (error) {
      console.error("CMP_AddMember_getMemberListColumnNames \n", error);
      this.setState({
        rejectMessage:
          stringConstants.CMPErrorMessage +
          ` while retrieving the ${stringConstants.MemberList} column data. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //Method to retrieve pending items from Member list
  private async getPendingItems() {
    try {
      this.setState({ showSpinner: true });
      //Getting the pending items from Member List
      let filterQuery = "Status eq '" + stringConstants.pendingStatus + "'";
      const sortColumn = "Created";
      const pendingItems: any[] = await commonServiceManager.getItemsSortedWithFilter(stringConstants.MemberList, filterQuery, sortColumn);
      this.setState({
        championList: pendingItems,
        showSpinner: false
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
      let updateResponse = await commonServiceManager.updateMultipleItems(stringConstants.MemberList, updateMemberObject, selectedIDs);

      //Refresh the items shown in the grid
      await this.getPendingItems();

      if (updateResponse) {
        //Updating state variables based on the approval action
        if (statusText === stringConstants.approvedStatus) {
          this.setState({
            approveMessage: LocaleStrings.ChampionApprovedMessage,
            selectedChampions: [],
            rejectMessage: "",
            isAllSelected: false,
            showSpinner: false
          });
        } else if (statusText === stringConstants.rejectedStatus) {
          this.setState({
            rejectMessage: LocaleStrings.ChampionRejectedMessage,
            selectedChampions: [],
            approveMessage: "",
            isAllSelected: false,
            showSpinner: false
          });
        }
      } else {
        //If any error occurs during batch update
        this.setState({
          rejectMessage: stringConstants.CMPErrorMessage + "while approving/rejecting champion request(s).",
          selectedChampions: [],
          isAllSelected: false,
          showSpinner: false
        });
      }
    }
    catch (error) {
      //Refresh the items shown in the grid
      await this.getPendingItems();

      this.setState({
        rejectMessage: stringConstants.CMPErrorMessage + "while approving/rejecting champion request(s). Below are the details: \n" + JSON.stringify(error),
        selectedChampions: [],
        isAllSelected: false,
        showSpinner: false
      });
      console.error("CMP_ApproveChampion_updatePendingItems \n", JSON.stringify(error));
    }
  }

  //Update all selected champions to new array
  public selectChampions(isChecked: boolean, key: number, selectAll: boolean) {
    //When "Select All" is checked
    if (selectAll && isChecked) {
      this.setState({ isAllSelected: true });
      let selectedChampions: any = [];
      this.state.filteredChampionList.forEach((event: ISPList) => {
        selectedChampions.push(event.ID);
      });
      this.setState({ selectedChampions: selectedChampions });
    }
    // When "Select All" is unchecked
    else if (selectAll && !isChecked) {
      this.setState({ isAllSelected: false, selectedChampions: [] });
    }
    else {
      //When checkbox is checked
      if (isChecked) {
        let selectedEvents = this.state.selectedChampions;
        selectedEvents.push(key);
        this.setState({ selectedChampions: selectedEvents });

        //Automatically check the "Select All" option when the last checkbox is checked
        if (selectedEvents.length === this.state.filteredChampionList.length) {
          this.setState({ isAllSelected: true });
        }

      }
      //When checkbox is unchecked
      else {
        const selectedEvents = this.state.selectedChampions.filter((eventId: any) => {
          return eventId !== key;
        });
        this.setState({
          isAllSelected: false,
          selectedChampions: selectedEvents
        });
      }
    }
  }

  //populate member list column display names into the states
  public populateColumnNames() {
    const enabledSettingsArray = this.state.configListSettings.filter((setting) => setting.Value === stringConstants.EnabledStatus);
    for (let setting of enabledSettingsArray) {
      const columnObject = this.state.memberListColumnNames.find((column) => column.InternalName === setting.Title);
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

  //Set pagination properties
  private pagination = paginationFactory({
    page: 1,
    sizePerPage: 10,
    lastPageText: '>>',
    firstPageText: '<<',
    nextPageText: '>',
    prePageText: '<',
    showTotal: true,
    alwaysShowAllBtns: false,
    paginationTotalRenderer: (from: any, to: any, size: any) => {
      const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} Results` : ""
      return (
        <span className="react-bootstrap-table-pagination-total" aria-live="polite" role="alert">
          &nbsp;{resultsFound}
        </span>
      )
    }
  });

  //Get Table Header Class
  private getTableHeaderClass(enabledColumnCount: number) {
    if (enabledColumnCount === 3)
      return styles.championsApprovalTableHeaderWithAllCols;
    else if (enabledColumnCount === 2)
      return styles.championsApprovalTableHeaderWithSixCols;
    else if (enabledColumnCount === 1)
      return styles.championsApprovalTableHeaderWithFiveCols;
    else if (enabledColumnCount === 0)
      return styles.championsApprovalTableHeaderWithFourCols;
  }

  //Get Table Body Class
  private getTableBodyClass(enabledColumnCount: number) {
    if (enabledColumnCount === 3)
      return styles.championsApprovalTableBodyWithAllCols;
    else if (enabledColumnCount === 2)
      return styles.championsApprovalTableBodyWithSixCols;
    else if (enabledColumnCount === 1)
      return styles.championsApprovalTableBodyWithFiveCols;
    else if (enabledColumnCount === 0)
      return styles.championsApprovalTableBodyWithFourCols;
  }

  //render the sort caret on the header column for accessbility
  customSortCaret = (order: any, column: any) => {
    if (!order) {
      return (
        <span className="sort-order">
          <span className="dropdown-caret">
          </span>
          <span className="dropup-caret">
          </span>
        </span>);
    }
    else if (order === 'asc') {
      return (
        <span className="sort-order">
          <span className="dropup-caret">
          </span>
        </span>);
    }
    else if (order === 'desc') {
      return (
        <span className="sort-order">
          <span className="dropdown-caret">
          </span>
        </span>);
    }
    return null;
  }


  // format the cell for Champion Name
  championFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
    return (
      <Person
        personQuery={gridRow.Title}
        view={3}
        personCardInteraction={1}
        className='champion-person-card'
      />
    );
  }

  public render() {
    //storing number of dropdowns got enabled
    const enabledColumnCount = (this.state.countryColumnName !== "" ? 1 : 0) +
      (this.state.regionColumnName !== "" ? 1 : 0) + (this.state.groupColumnName !== "" ? 1 : 0);
    const championsTableHeader: any = [
      {
        dataField: "ID",
        headerFormatter: () => {
          return (
            <Checkbox
              onChange={(_eve: any, isChecked: boolean) => {
                this.selectChampions(isChecked, -1, true);
              }}
              className={styles.selectAllCheckbox}
              checked={this.state.isAllSelected}
              ariaLabel={LocaleStrings.SelectAllChampions}
              disabled={this.state.showSpinner || this.state.championList.length === 0}
            />
          );
        },
        headerTitle: () => LocaleStrings.SelectAllChampions,
        title: () => LocaleStrings.SelectChampion,
        attrs: (_cell: any, row: any) => ({ key: row.ID }),
        formatter: (_: any, gridRow: any) => {
          return (
            <Checkbox
              onChange={(_eve: any, isChecked: boolean) => {
                this.selectChampions(isChecked, gridRow.ID, false);
              }}
              className={styles.selectItemCheckbox}
              checked={this.state.selectedChampions.includes(gridRow.ID)}
              ariaLabel={LocaleStrings.SelectChampion}
              disabled={this.state.showSpinner}
            />
          );
        },
        searchable: false
      },
      {
        dataField: "FirstName",
        text: LocaleStrings.PeopleNameGridHeader,
        headerTitle: true,
        formatter: this.championFormatter,
        searchable: true,
        sort: true,
        sortCaret: this.customSortCaret
      },
      {
        dataField: "Title",
        text: LocaleStrings.EmailLabel,
        headerTitle: true,
        title: true,
        searchable: true,
        sort: true,
        sortCaret: this.customSortCaret
      },
      {
        dataField: "Region",
        text: this.state.regionColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.regionColumnName === ""
      },
      {
        dataField: "Country",
        text: this.state.countryColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.countryColumnName === ""
      },
      {
        dataField: "Group",
        text: this.state.groupColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.groupColumnName === ""
      },
      {
        dataField: "FocusArea",
        text: LocaleStrings.FocusAreaGridHeader,
        headerTitle: true,
        title: true,
        searchable: false
      }
    ];
    return (
      <div className={styles.approvalsContainer}>
        {this.state.approveMessage &&
          <Label className={styles.approveMessage + ' col-xl-5 col-lg-5 col-md-6 col-sm-8 col-xs-9'} aria-live="polite" role="alert">
            <img
              src={require('../assets/TOTImages/tickIcon.png')}
              alt={LocaleStrings.SuccessIcon}
              className={styles.tickImage}
            />
            {this.state.approveMessage}
          </Label>
        }
        {this.state.rejectMessage &&
          <Label className={styles.rejectMessage + ' col-xl-5 col-lg-5 col-md-6 col-sm-8 col-xs-9'} aria-live="polite" role="alert">
            {this.state.rejectMessage}
          </Label>
        }
        <ToolkitProvider
          bootstrap4
          keyField="ID"
          data={this.state.championList}
          columns={championsTableHeader}
          search={
            {
              afterSearch: (newResult: ISPList[]) => {
                this.setState({
                  filteredChampionList: newResult,
                  isAllSelected: newResult.length === this.state.selectedChampions.length ? true : false
                });
              }
            }
          }
        >
          {
            (props: ToolkitContextType) => (
              <div>
                {this.state.championList.length > 0 &&
                  <>
                    <div className={'col-xl-5 col-lg-5 col-md-6 col-sm-7 col-xs-9' + " " + styles.searchboxPadding}>
                      <SearchBox
                        placeholder={LocaleStrings.ApproveChampionSearchboxPlaceholder}
                        onChange={(_, searchedText) => props.searchProps.onSearch(searchedText)}
                        className={styles.approvalsSearchbox}
                      />
                    </div>
                    {this.state.selectedChampions.length > 0 &&
                      <Label className={styles.selectedRowText}>
                        {this.state.selectedChampions.length} {LocaleStrings.ChampionsSelectedLabel}
                      </Label>
                    }
                  </>
                }
                <div>
                  <BootstrapTable
                    striped
                    {...props.baseProps}
                    table-responsive={true}
                    pagination={this.pagination}
                    wrapperClasses={styles.approvalsTableWrapper}
                    headerClasses={this.getTableHeaderClass(enabledColumnCount)}
                    bodyClasses={this.getTableBodyClass(enabledColumnCount)}
                    noDataIndication={() => (
                      <div className={styles.noRecordsArea}>
                        {this.state.showSpinner ?
                          <Spinner
                            label={LocaleStrings.ProcessingSpinnerLabel}
                            size={SpinnerSize.large}
                          /> :
                          <>
                            <img
                              src={require('../assets/CMPImages/Norecordsicon.svg')}
                              alt={LocaleStrings.NoRecordsIcon}
                              className={styles.noRecordsImg}
                              aria-hidden={true}
                            />
                            <span className={styles.noRecordsLabels} aria-live='polite' role="alert" tabIndex={0}>
                              {this.state.championList.length === 0 ?
                                LocaleStrings.NoChampionsMessage
                                :
                                LocaleStrings.NoSearchResults
                              }
                            </span>
                          </>
                        }
                      </div>
                    )}
                  />
                </div>
              </div>
            )
          }
        </ToolkitProvider>
        {this.state.showSpinner && this.state.championList.length > 0 &&
          <Spinner
            label={LocaleStrings.ProcessingSpinnerLabel}
            size={SpinnerSize.large}
          />
        }
        {this.state.championList.length > 0 &&
          <div className={styles.manageApprovalsBtnArea}>
            <button
              className={`btn ${styles.approveBtn}`}
              onClick={e => this.updatePendingItems(stringConstants.approvedStatus, this.state.selectedChampions)}
              title={LocaleStrings.ApproveButton}
              disabled={this.state.selectedChampions.length === 0}
            >
              <Icon iconName="Completed" className={styles.approveBtnIcon} />
              <span className={styles.approveBtnLabel}>{LocaleStrings.ApproveButton}</span>
            </button>
            <button
              className={"btn " + styles.rejectBtn}
              onClick={e => this.updatePendingItems(stringConstants.rejectedStatus, this.state.selectedChampions)}
              title={LocaleStrings.RejectButton}
              disabled={this.state.selectedChampions.length === 0}
            >
              <Icon iconName="ErrorBadge" className={styles.rejectBtnIcon} />
              <span className={styles.rejectBtnLabel}>{LocaleStrings.RejectButton}</span>
            </button>
          </div>
        }
      </div>
    );
  }
}

export default ApproveChampion;
