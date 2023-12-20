import {
  ISPHttpClientOptions, SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as _ from "lodash";
import { Icon, initializeIcons } from "office-ui-fabric-react";
import { Dropdown, IDropdownOption, IDropdownStyles } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Dialog, DialogType } from '@fluentui/react/lib/Dialog';
import * as React from "react";
import Button from "react-bootstrap/esm/Button";
import commonServices from "../Common/CommonServices";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";
import "../scss/Championleaderboard.scss";
import { IConfigList } from './ManageConfigSettings';
import RecordEvents from "./RecordEvents";
import ChampionEvents from "./ChampionEvents";
import EventsChart from "./EventsChart";
import Row from "react-bootstrap/esm/Row";
import Col from "react-bootstrap/esm/Col";
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

initializeIcons();

//global variables
let commonServiceManager: commonServices;
export interface ISidebarStateProps {
  becomec: boolean;
  context?: any;
  onClickCancel: () => void; //will redirects to back/home
  callBack?: Function;
  siteUrl: string;
  setEventsSubmissionMessage?: Function;
  currentThemeName?: string;
}
interface IUserDetail {
  ID: number;
  LoginName: string;
}
export class ISPLists {
  public value: ISPList[];
}
export class ISPList {
  public Title: string;
  public FirstName: string;
  public LastName: string;
  public Country: string;
  public Status: string;
  public Role: string;
  public Region: string;
  public Points: number;
  public Group: string;
  public FocusArea: string;
}
interface IState {
  currentUser: ISPList;
  UserDetails: Array<any>;
  bc: boolean;
  isLoaded: boolean;
  form: boolean;
  totalUserPointsfromList: number;
  totalUsers: number;
  userRank: number;
  isActive: boolean;
  countries: Array<any>;
  regions: Array<any>;
  roles: Array<any>;
  status: Array<any>;
  memberData: any;
  selectedFocusAreas: any;
  multiSelectChoices: any;
  buttonText: any;
  isMember: boolean;
  emailValue: string;
  sitename: string;
  inclusionpath: string;
  edetails: Array<string>;
  edetailsIds: Array<EventList>;
  configListSettings: Array<IConfigList>;
  memberListColumnNames: Array<any>;
  regionColumnName: string;
  countryColumnName: string;
  groupColumnName: string;
  showRecordEventPopup: boolean;
  showDashBoardPopup: boolean;
  memberEvents: Array<any>;
  selectedMemberID: string;
  isDesktop: boolean;
}
export interface EventList {
  Title: string;
  Id: number;
}

let displayName: string = "";

export default class Sidebar extends React.Component<ISidebarStateProps, IState> {
  constructor(props: ISidebarStateProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });

    this.state = {
      bc: this.props.becomec,
      form: false,
      isLoaded: false,
      UserDetails: [],
      currentUser: new ISPList(),
      totalUserPointsfromList: 0,
      isActive: false,
      totalUsers: 0,
      userRank: 0,
      countries: [],
      regions: [],
      roles: [],
      status: [],
      memberData: { region: "", role: "", status: "", country: "" },
      selectedFocusAreas: [],
      multiSelectChoices: [],
      buttonText: LocaleStrings.BecomeChampionLabel,
      isMember: false,
      emailValue: "",
      sitename: siteconfig.sitename, //getting from siteconfig
      inclusionpath: siteconfig.inclusionPath, //getting from siteconfig
      edetails: [],
      edetailsIds: [],
      configListSettings: [],
      memberListColumnNames: [],
      regionColumnName: "",
      countryColumnName: "",
      groupColumnName: "",
      showRecordEventPopup: false,
      showDashBoardPopup: false,
      memberEvents: [],
      selectedMemberID: "",
      isDesktop: true
    };
    this.handleInput = this.handleInput.bind(this);
    this._createorupdateItem = this._createorupdateItem.bind(this);
    this._getPeoplePickerItems = this._getPeoplePickerItems.bind(this);
    this._getListData = this._getListData.bind(this);
    this.optionsEventsList = this.optionsEventsList.bind(this);
    this.onFocusAreaChange = this.onFocusAreaChange.bind(this);
    this.populateColumnNames = this.populateColumnNames.bind(this);
    this.updateRecordEventsPopupState = this.updateRecordEventsPopupState.bind(this);

    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );

  }

  //getting members details from membelist with all columns
  public options = (optionArray: any) => {
    let myoptions = [];
    if (optionArray !== undefined) {
      myoptions.push({ key: "All", text: "All" });
      optionArray.forEach((element: any) => {
        myoptions.push({ key: element, text: element });
      });
    }
    return myoptions;
  }

  public optionsEventsList() {
    let optionArray: any = [];
    let optionArrayIds: any = [];
    if (this.state.edetails.length == 0)
      this.props.context.spHttpClient
        .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Events List')/Items", SPHttpClient.configurations.v1)
        .then(async (response: SPHttpClientResponse) => {
          if (response.status === 200) {
            await response.json().then((responseJSON: any) => {
              let i = 0;
              while (i < responseJSON.value.length) {
                if (responseJSON.value[i] && responseJSON.value[i].hasOwnProperty("Title") && responseJSON.value[i].IsActive) {
                  optionArray.push(responseJSON.value[i].Title);
                  optionArrayIds.push({
                    Title: responseJSON.value[i].Title,
                    Id: responseJSON.value[i].Id,
                  });
                }
                i++;
              }
              this.setState({
                edetails: optionArray,
                edetailsIds: optionArrayIds,
              });
            });
          }
        })
        .catch(() => {
          throw new Error("Asynchronous error");
        });

    //getting event details from membelist
    let myOptions = [];
    myOptions.push({ key: "Select Event Type", text: "Select Event Type" });
    this.state.edetails.forEach((element: any) => {
      myOptions.push({ key: element, text: element });
    });
    return myOptions;
  }

  public dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: "auto" },
  };

  public handleInput(event: any, key: string) {
    let user: any = this.state.currentUser;
    user[key] = event.target.value;
    this.setState({ currentUser: user });
  }

  //getting Region and Country details from membelist
  public componentDidMount() {
    this.optionsEventsList();
    this.props.context.spHttpClient
      .get(
        "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Region')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((regions: any) => {
          if (!regions.error) {
            this.props.context.spHttpClient
              .get(
                "/" +
                this.state.inclusionpath +
                "/" +
                this.state.sitename +
                "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Country')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line:no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((countries: any) => {
                  if (!countries.error) {
                    this.setState({
                      regions: regions.Choices,
                      countries: countries.Choices,
                    });
                  }
                });
              });
          }
        });
      });

    this.props.context.spHttpClient
      .get(
        "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Group')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((roles: any) => {
          if (!roles.error) {
            this.props.context.spHttpClient
              .get(
                "/" +
                this.state.inclusionpath +
                "/" +
                this.state.sitename +
                "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('FocusArea')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line: no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((status: any) => {
                  if (!status.error) {
                    this.setState({
                      roles: roles.Choices,
                      status: status.Choices,
                    });
                  }
                });
              });
          }
        });
      });
    // Adding window resize event listener while mounting the component
    window.addEventListener("resize", this.resize.bind(this));
    this.resize();
    //set css properties to Person card control
    this.updatePersonCardCSS();
  }

  // Set the state object for screen size
  resize = () => this.setState({ isDesktop: window.innerWidth > 568 });

  //set css properties to Person card control
  public updatePersonCardCSS() {
    setTimeout(() => {
      const sidebarPersonWrapper = document.getElementById("sidebar-person-wrapper")?.querySelector("mgt-person")
        ?.shadowRoot?.querySelector("mgt-flyout")?.querySelector(".vertical");
      sidebarPersonWrapper?.setAttribute("style", "row-gap:10px;");
      if (this.state.isDesktop) {
        sidebarPersonWrapper?.querySelector(".avatar-wrapper")?.setAttribute("style", "width: 100px; height: 100px;");
        sidebarPersonWrapper?.querySelector(".details-wrapper")?.querySelector(".line1")?.setAttribute("style",
          "width: 180px;overflow-wrap: break-word;text-align: center;");
      }
      else {
        sidebarPersonWrapper?.querySelector(".avatar-wrapper")?.setAttribute("style", "");
        sidebarPersonWrapper?.querySelector(".details-wrapper")?.querySelector(".line1")
          ?.setAttribute("style", "font-size:14px;width: 140px;overflow-wrap: break-word;text-align: center;");
      }
    }, 5000);
  }

  public async componentDidUpdate(prevProps: Readonly<ISidebarStateProps>, prevState: Readonly<IState>, snapshot?: any) {
    if (prevProps != this.props) {
      this.componentWillMount();
    }
    if (prevState.multiSelectChoices !== this.state.multiSelectChoices) {
      this.setState({
        selectedFocusAreas: this.state.multiSelectChoices
      });
    }
    //Remove "All" from the array to store it in Members List.
    if (prevState.selectedFocusAreas !== this.state.selectedFocusAreas) {
      let idx = this.state.selectedFocusAreas.indexOf(stringsConstants.AllLabel);
      if (idx != -1)
        this.state.selectedFocusAreas.splice(idx, 1);
    }

    if (prevState.form !== this.state.form || prevState.isActive !== this.state.isActive) {
      if (this.state.form && this.state.isActive) {
        this.setState({
          configListSettings: [],
          memberListColumnNames: [],
          regionColumnName: "",
          countryColumnName: "",
          groupColumnName: ""
        });
        await this.getConfigListSettings();
        await this.getMemberListColumnNames();
      }
    }

    //update column states with member list column display names 
    if (prevState.configListSettings !== this.state.configListSettings ||
      prevState.memberListColumnNames !== this.state.memberListColumnNames) {
      if (this.state.configListSettings.length > 0 && this.state.memberListColumnNames.length > 0)
        this.populateColumnNames();
    }

    if (prevState.isDesktop !== this.state.isDesktop) {
      //set css properties to Person card control
      this.updatePersonCardCSS();
    }
  }

  // Before unmounting, remove event listener
  componentWillUnmount() {
    window.removeEventListener("resize", this.resize.bind(this));
  }

  //Get current user's details from Member list and Event track details to display rank and points on side bar
  public componentWillMount() {
    this.props.context.spHttpClient
      .get(

        "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          if (!datauser.error) {
            this.props.context.spHttpClient
              .get(
                "/" +
                this.state.inclusionpath +
                "/" +
                this.state.sitename +
                "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000",
                SPHttpClient.configurations.v1
              )
              .then((response: SPHttpClientResponse) => {
                response.json().then((datada: any) => {
                  let memberDataIds = datada.value.find(
                    (d: { Title: string }) =>
                      d.Title.toLowerCase() === datauser.Email.toLowerCase()
                  );
                  let memberData =
                    memberDataIds !== undefined ? memberDataIds.ID : 0;
                  this.setState({
                    selectedMemberID: memberData
                  });
                  if (memberData === 0)
                    this.setState({
                      emailValue: datauser.Email,
                      isMember: false,
                      buttonText: LocaleStrings.BecomeChampionLabel, //if employee want to become a champion enabling this button
                      isLoaded: true,
                    });
                  else
                    this.setState({
                      emailValue: datauser.Email,
                      isMember: true,
                      buttonText: LocaleStrings.ChampionSubmissionPendingLabel, //if employee already raised for a champion enabling this button
                      isLoaded: true,
                    });
                  localStorage.setItem("memberid", memberData); // storing memberid in local storage

                  //get first name and last name from the user profile properties
                  let firstName = "";
                  let lastName = "";
                  for (let i = 0; i < datauser.UserProfileProperties.length; i++) {
                    if (firstName === "" || lastName === "") {
                      if (datauser.UserProfileProperties[i].Key === "FirstName") {
                        firstName = datauser.UserProfileProperties[i].Value;
                      }
                      if (datauser.UserProfileProperties[i].Key === "LastName") {
                        lastName = datauser.UserProfileProperties[i].Value;
                      }
                    }
                    else {
                      break;
                    }
                  }
                  //based on user role (champion or manager) then we are showing champion details
                  let user = this.state.currentUser;
                  user["FirstName"] = firstName;
                  user["LastName"] = lastName;
                  user["Title"] = datauser.Email;
                  user["Country"] = datauser.Country;
                  user["Region"] = datauser.Region;
                  user["Group"] = datauser.Group;
                  user["FocusArea"] = datauser.FocusArea;
                  displayName = firstName + " " + lastName;

                  this.setState({ currentUser: user });
                  if (!datada.error) {
                    let totalchamps: number = 0;
                    totalchamps = datada.value.filter((x: any) =>
                      (x.Role.toLowerCase() === "champion" ||
                        x.Role.toLowerCase() === "manager") &&
                        x.Status !== null &&
                        x.Status !== undefined
                        ? x.Status.toLowerCase() === "approved"
                        : false
                    ).length;
                    if (
                      this.state.isMember === true &&
                      (memberDataIds.Role == "Champion" ||
                        memberDataIds.Role === "Manager") &&
                      memberDataIds.Status === "Approved"
                    )
                      this.props.context.spHttpClient
                        .get(
                          "/" +
                          this.state.inclusionpath +
                          "/" +
                          this.state.sitename +
                          "/_api/web/lists/GetByTitle('Event Track Details')/Items?$filter= Status eq 'Approved' or Status eq null or Status eq ''&$top=5000",
                          SPHttpClient.configurations.v1
                        )
                        .then((responseeventsdetails: SPHttpClientResponse) => {
                          responseeventsdetails
                            .json()
                            .then((eventsdatauser: any) => {
                              this.setState({
                                memberEvents: eventsdatauser.value
                              });
                              if (!eventsdatauser.error) {
                                let memberids: any = _.uniqBy(
                                  eventsdatauser.value,
                                  "MemberId"
                                );
                                let memcount: Array<any> = [];
                                for (let i = 0; i < memberids.length; i++) {
                                  if (datada.value.findIndex((v: { ID: any }) => v.ID === memberids[i].MemberId) !== -1) {
                                    let totalUserPoints = 0;
                                    eventsdatauser.value.filter((z: any) => z.MemberId === memberids[i].MemberId)
                                      .map((z: any) => { totalUserPoints = totalUserPoints + z.Count; });

                                    memcount.push({
                                      id: memberids[i].MemberId,
                                      points: totalUserPoints,
                                    });
                                  }
                                }

                                //Assign zero points and get data of the approved members who hasn't participated in any event
                                let tempArray: any = [];
                                for (let i = 0; i < datada.value.length; i++) {
                                  if (memcount.findIndex((member: { id: any }) => member.id === datada.value[i].ID) === -1 &&
                                    datada.value[i].Status === stringsConstants.approvedStatus
                                  ) {
                                    tempArray.push({
                                      id: datada.value[i].ID,
                                      points: 0,
                                    });
                                  }
                                }
                                memcount = [...memcount, ...tempArray];
                                let pointsTotal = 0;
                                let rank: number;

                                //Sorting
                                //Intially sort approved champions in the descending order of points count 
                                //if champion doesn't have any points sort them in ascending order of their ids
                                memcount
                                  .sort((a, b) => {

                                    if (a.points < b.points) return 1;

                                    if (a.points > b.points) return -1;

                                    if (a.id > b.id) return 1;

                                    if (a.id < b.id) return -1;

                                  })
                                  .map((x: any, ind: number) => {
                                    if (
                                      x.id ===
                                      datada.value.find(
                                        (d: { Title: string }) =>
                                          d.Title.toLowerCase() ===
                                          datauser.Email.toLowerCase()
                                      ).ID
                                    ) {
                                      rank = ind + 1;
                                      pointsTotal = x.points;
                                    }
                                  });
                                this.setState({
                                  isLoaded: true,
                                  totalUserPointsfromList: pointsTotal,
                                  totalUsers: totalchamps,
                                  userRank: rank,
                                });
                              }
                            });
                        });
                  }
                });
              });
          }
        });
      });
  }

  //Get settings from config list
  private async getConfigListSettings() {
    try {
      const configListData: IConfigList[] = await commonServiceManager.getMemberListColumnConfigSettings();
      if (configListData.length === 3) {
        this.setState({ configListSettings: configListData });
      }
      else {
        console.log(
          stringsConstants.CMPErrorMessage +
          ` while loading the page. There could be a problem with the ${stringsConstants.ConfigList} data.`
        );
      }
    }
    catch (error) {
      console.error("CMP_Sidebar_getConfigListSettings \n", error);
      console.log(
        stringsConstants.CMPErrorMessage +
        `while retrieving the ${stringsConstants.ConfigList} settings. Below are the details: \n` +
        JSON.stringify(error),
      );
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
      console.error("CMP_Sidebar_getMemberListColumnNames \n", error);
      console.log(
        stringsConstants.CMPErrorMessage +
        ` while retrieving the ${stringsConstants.MemberList} column data. Below are the details: \n` +
        JSON.stringify(error),
      );
    }
  }

  //populate member list column display names into the states
  private populateColumnNames() {
    const enabledSettingsArray = this.state.configListSettings.filter((setting) => setting.Value === stringsConstants.EnabledStatus);
    for (let setting of enabledSettingsArray) {
      const columnObject = this.state.memberListColumnNames.find((column) => column.InternalName === setting.Title);
      if (columnObject.InternalName === stringsConstants.RegionColumn) {
        this.setState({ regionColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringsConstants.CountryColumn) {
        this.setState({ countryColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringsConstants.GroupColumn) {
        this.setState({ groupColumnName: columnObject.Title });
      }
    }
  }

  //getting extra symbol, so using default menthod
  public onRenderCaretDown = (): JSX.Element => {
    return <span></span>;
  }

  //when user raised for become a champion then his state would be champion and pending
  public async _createorupdateItem() {
    let usersave = this.state.currentUser;
    usersave.Country = this.state.memberData.country;
    usersave.Region = this.state.memberData.region;
    usersave.FocusArea = this.state.selectedFocusAreas.length > 0 ? this.state.selectedFocusAreas : [stringsConstants.TeamWorkLabel];
    usersave.Group = this.state.memberData.group;
    usersave.Role = "Champion";
    usersave.Status = "Pending";

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(usersave),
    };

    let flag = await this._getListData(usersave.Title);
    if (flag == 0) {
      const url: string = "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/items";
      if (this.props.context)
        this.props.context.spHttpClient
          .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
          .then((response: SPHttpClientResponse) => {
            if (response.status === 201) {
              alert(LocaleStrings.ChampionRequestSubmitSuccessMessage);
              {
                this.props.onClickCancel();
              }
            } else {
              alert(
                "Response status " +
                response.status +
                " - " +
                response.statusText
              );
            }
          });
    } else {
      alert("Record Already Exist");
      {
        this.props.onClickCancel();
      }
    }
    this.setState({
      currentUser: new ISPList(),
    });
    this.setState({ isActive: !this.state.isActive, form: !this.state.form });
  }

  //Get current user's details from Member List
  private async _getListData(email: any): Promise<any> {
    return this.props.context.spHttpClient
      .get("/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('Member List')/Items?$filter=Title eq '" + email.toLowerCase() + "'",
        SPHttpClient.configurations.v1
      )
      .then(async (response: SPHttpClientResponse) => {
        if (response.status === 200) {
          let flag = 0;
          await response.json().then((responseJSON: any) => {
            let i = 0;
            while (i < responseJSON.value.length) {
              if (
                responseJSON.value[i].Title.toLowerCase() == email.toLowerCase()
              ) {
                flag = 1;
                return flag;
              }
              i++;
            }
            return flag;
          });
          return flag;
        }
      });
  }

  public addDefaultSrc(ev: any) {
    ev.target.src = require("../assets/images/noprofile.png"); //if no profile then we are showing default image
  }

  private _getPeoplePickerItems(items: any[]) {
    let userarr: IUserDetail[] = [];
    items.forEach((user) => {
      userarr.push({ ID: user.id, LoginName: user.loginName });
      user["FirstName"] = user.text.split(",")[0] || "";
      user["LastName"] = user.text.split(",")[1] || "";
      user["Title"] = user.loginName.split("|")[2] || "";
      this.setState({ currentUser: user });
    });
    this.setState({ UserDetails: userarr });
  }

  public filterUsers(type: string, value: any) {
    if (value.target.innerText !== "All") {
      this.setState({
        memberData: {
          ...this.state.memberData,
          [type]: value.target.innerText,
        },
      });
    }
  }

  //Set state variable whenever the Focus Area dropdown is changed
  public onFocusAreaChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
    if (item === undefined) {
      return;
    }
    //Select all the dropdown options when "All" is selected.
    if (item.key === stringsConstants.AllLabel && item.selected) {
      this.setState({
        multiSelectChoices: this.options(this.state.status).map((option) => option.key as string)
      });
    } //Clear all the dropdown options when "All" is unselected
    else if (item.key === stringsConstants.AllLabel) {
      this.setState({ multiSelectChoices: [] });
    } //When an option selected from the dropdown other than "All"
    else if (item.selected) {
      const newKeys = [item.key as string];
      if (this.state.multiSelectChoices.length === this.state.status.length - 1) {
        newKeys.push(stringsConstants.AllLabel);
      }
      this.setState({ multiSelectChoices: [...this.state.multiSelectChoices, ...newKeys] });
    } //When an option unselected from the dropdown other than "All"
    else {
      this.setState({
        multiSelectChoices: this.state.multiSelectChoices.filter((key: any) => key !== item.key && key !== stringsConstants.AllLabel)
      });
    }
  }

  private updateRecordEventsPopupState(show: boolean) {
    this.setState({ showRecordEventPopup: show });
  }

  public render() {
    const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
    return (
      <div className="Championleaderboard">
        {this.state.isLoaded && (
          <div className="sidenav">
            <div className="imagePointsArea">
              <div id="sidebar-person-wrapper">
                <Person
                  personQuery="me"
                  view={3}
                  personCardInteraction={1}
                  verticalLayout={true}
                  className="championSideBar"
                />
              </div>
              {!this.state.bc && !this.state.form && (
                <>
                  <div className="links-wrapper">
                    <div className="sidebar-action-link"
                      title={LocaleStrings.DashboardLabel}
                      onClick={() => this.setState({ showDashBoardPopup: true })}
                    >
                      <img
                        src={require("../assets/CMPImages/Dashboard.svg")}
                        alt="dashboard"
                        className="action-img"
                      />
                      <span className="action-text">{LocaleStrings.DashboardLabel}</span>
                    </div>
                    <div className="sidebar-action-link"
                      title={LocaleStrings.RecordEventLabel}
                      onClick={() => this.setState({ showRecordEventPopup: true })}
                    >
                      <img
                        src={require("../assets/CMPImages/RecordEvents.svg")}
                        alt="record events"
                        className="action-img"
                      />
                      <span className="action-text">{LocaleStrings.RecordEventLabel}</span>
                    </div>
                  </div>
                  <div>
                    {/* here we are showing rank and points  */}
                    <div className="pointcircle">
                      <div className="insidecircle">
                        <div className="pointsscale">
                          <div><Icon iconName="FavoriteStarFill" id="star" className="yellowStar" /></div>
                          <div className="pointsValueLabel">{this.state.totalUserPointsfromList} {LocaleStrings.CMPSideBarPointsLabel}</div>
                        </div>
                        <div className="line" />
                        <div className="globalrank">
                          <div>
                            <span className="bold">{this.state.userRank} </span>
                            {LocaleStrings.CMPSideBarGlobalRankLabel}
                          </div>
                          <div>
                            of{" "}
                            {this.state.totalUsers + " "}
                            {LocaleStrings.CMPSideBarChampionsLabel}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}
              {this.state.bc && (
                <div>
                  {
                    <Button
                      variant="primary"
                      className="bc-btn"
                      disabled={this.state.isMember ? true : false}
                      onClick={() =>
                        this.setState({
                          form: !this.state.form,
                          isActive: !this.state.isActive,
                        })
                      }
                    >
                      <span> {this.state.buttonText}</span>
                    </Button>
                  }
                </div>
              )}
            </div>
            {this.state.form && this.state.isActive && (
              // become a champion form
              <div className="bc-form-main-area">
                <div className="bc-form">
                  <div className="bc-form-close-icon-area">
                    <Icon
                      iconName="ChromeClose"
                      className="bc-form-close-icon"
                      onClick={() => {
                        this.setState({
                          form: !this.state.form,
                          isActive: !this.state.isActive
                        });
                      }}
                    />
                  </div>
                  <label htmlFor="fname" className="bc-label">
                    {LocaleStrings.FirstNameLabel}
                  </label>
                  <TextField
                    value={
                      this.state.currentUser.FirstName
                        ? this.state.currentUser.FirstName
                        : ""
                    }
                    onChange={(evt) => this.handleInput(evt, "FirstName")}
                  />
                  <label htmlFor="lname" className="bc-label">
                    {LocaleStrings.LastNameLabel}
                  </label>
                  <TextField
                    value={
                      this.state.currentUser.LastName
                        ? this.state.currentUser.LastName
                        : ""
                    }
                    onChange={(evt) => this.handleInput(evt, "LastName")}
                  />
                  <label htmlFor="email" className="bc-label">
                    {LocaleStrings.EmailIDLabel}
                  </label>
                  <TextField
                    value={
                      this.state.currentUser.Title
                        ? this.state.currentUser.Title
                        : ""
                    }
                    onChange={(evt) => this.handleInput(evt, "Title")}
                  />
                  {this.state.regionColumnName !== "" &&
                    <>
                      <label htmlFor="Region" className="bc-label" title={this.state.regionColumnName}>
                        {this.state.regionColumnName}
                      </label>
                      <Dropdown
                        onChange={(event: any) => this.filterUsers("region", event)}
                        placeholder={"Select " + this.state.regionColumnName}
                        options={this.options(this.state.regions)}
                        styles={this.dropdownStyles}
                        onRenderCaretDown={this.onRenderCaretDown}
                        defaultValue={this.state.currentUser.Region}
                        calloutProps={{ className: "nonMemberDdCallout" }}
                      />
                    </>
                  }
                  {this.state.countryColumnName !== "" &&
                    <>
                      <label htmlFor="Country" className="bc-label" title={this.state.countryColumnName}>
                        {this.state.countryColumnName}
                      </label>
                      <Dropdown
                        onChange={(event: any) =>
                          this.filterUsers("country", event)
                        }
                        placeholder={"Select " + this.state.countryColumnName}
                        options={this.options(this.state.countries)}
                        styles={this.dropdownStyles}
                        onRenderCaretDown={this.onRenderCaretDown}
                        defaultValue={this.state.currentUser.Country}
                        calloutProps={{ className: "nonMemberDdCallout" }}
                      />
                    </>
                  }
                  {this.state.groupColumnName !== "" &&
                    <>
                      <label htmlFor="Group" className="bc-label" title={this.state.groupColumnName}>
                        {this.state.groupColumnName}
                      </label>
                      <Dropdown
                        onChange={(event: any) => this.filterUsers("group", event)}
                        placeholder={"Select " + this.state.groupColumnName}
                        options={this.options(this.state.roles)}
                        styles={this.dropdownStyles}
                        onRenderCaretDown={this.onRenderCaretDown}
                        defaultValue={this.state.currentUser.Group}
                        calloutProps={{ className: "nonMemberDdCallout" }}
                      />
                    </>
                  }
                  <label htmlFor="Focus Area" className="bc-label" title={LocaleStrings.FocusAreaGridHeader}>
                    {LocaleStrings.FocusAreaGridHeader}
                  </label>
                  <Dropdown
                    onChange={this.onFocusAreaChange.bind(this)}
                    placeholder={LocaleStrings.FocusAreaPlaceholder}
                    options={this.options(this.state.status)}
                    styles={this.dropdownStyles}
                    onRenderCaretDown={this.onRenderCaretDown}
                    defaultValue={this.state.currentUser.FocusArea}
                    multiSelect
                    selectedKeys={this.state.multiSelectChoices}
                    calloutProps={{ className: "nonMemberDdCallout" }}
                  />
                  <Button
                    className="sub-btn"
                    type="reset"
                    onClick={() => this._createorupdateItem()}
                  >
                    {LocaleStrings.SubmitButton}
                  </Button>
                </div>
              </div>
            )}
            {this.state.showRecordEventPopup &&
              <RecordEvents
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                showRecordEventPopup={this.state.showRecordEventPopup}
                callBack={this.props.callBack}
                updateRecordEventsPopupState={this.updateRecordEventsPopupState}
                setEventsSubmissionMessage={this.props.setEventsSubmissionMessage}
                currentThemeName={this.props.currentThemeName}
              />
            }
            {this.state.showDashBoardPopup &&
              <Dialog
                hidden={!this.state.showDashBoardPopup}
                onDismiss={() => this.setState({ showDashBoardPopup: false })}
                modalProps={{
                  isBlocking: true,
                  className: `clb-dashboard-popup${isDarkOrContrastTheme ? " clb-dashboard-popup-" + this.props.currentThemeName : ""}`
                }}
                dialogContentProps={{ type: DialogType.normal, title: LocaleStrings.DashboardLabel, className: "clb-dialog-content" }}
              >
                <Row xl={2} lg={2} md={1} sm={1}>
                  <Col xl={5} lg={5} md={12} sm={12}>
                    <ChampionEvents
                      context={this.props.context}
                      filteredAllEvents={this.state.memberEvents}
                      parentComponent={stringsConstants.SidebarLabel}
                      selectedMemberID={this.state.selectedMemberID}
                    />
                  </Col>
                  <Col xl={7} lg={7} md={12} sm={12}>
                    <EventsChart
                      siteUrl={this.props.siteUrl}
                      context={this.props.context}
                      filteredAllEvents={this.state.memberEvents}
                      parentComponent={stringsConstants.SidebarLabel}
                      selectedMemberID={this.state.selectedMemberID}
                      currentThemeName={this.props.currentThemeName}
                    />
                  </Col>
                </Row>
              </Dialog>
            }
          </div >
        )
        }
      </div>
    );
  }
}
