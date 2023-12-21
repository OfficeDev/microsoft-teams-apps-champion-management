import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { app } from '@microsoft/teams-js-v2';
import { Component } from 'react';
import Accordion from 'react-bootstrap/Accordion';
import Card from 'react-bootstrap/Card';
import Col from 'react-bootstrap/esm/Col';
import Row from 'react-bootstrap/Row';
import Table from 'react-bootstrap/Table';
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { Dialog } from '@fluentui/react/lib/Dialog';
import { Icon, initializeIcons } from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import * as constants from '../constants/strings';
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import siteconfig from '../config/siteconfig.json';
import '../scss/Champions.scss';
import * as _ from "lodash";
import commonServices from '../Common/CommonServices';
import ChampionEvents from './ChampionEvents';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

initializeIcons();

let commonServiceManager: commonServices;
let currentUserName: string;

interface ChampionsCardsProps {
  type: string;
  events?: any;
  context: WebPartContext;
  siteUrl: string;
  callBack?: Function;
  loggedinUserEmail?: string;
  currentThemeName?: string;
}
interface ChampionsCardsState {
  isLoaded: boolean;
  loadCards: number;
  focusAreas: Array<any>;
  selectedFocusArea: string | number;
  search: string;
  filteredUsers: any;
  sitename: string;
  inclusionpath: string;
  users: any;
  memberEvents: Array<any>;
  showUserActivities: boolean;
  userActivities: Array<any>;
  filteredUserActivities: Array<any>;
  userActivitiesPerPage: number;
  pageNumber: number;
  selectedMemberDetails: any;
  events: Array<any>;
  selectedMemberID: string;
  isExpanded: boolean;
  cardPersonRefs: any;
}

export default class ChampionsCards extends Component<ChampionsCardsProps, ChampionsCardsState> {
  private leaderboardFocusAreaComboboxRef: React.RefObject<HTMLDivElement>;
  private mainComboboxRef: React.RefObject<IComboBox>;
  constructor(props: any) {
    super(props);
    this.leaderboardFocusAreaComboboxRef = React.createRef();
    this.mainComboboxRef = React.createRef();
    this.state = {
      isLoaded: false,
      loadCards: 0,
      focusAreas: [],
      selectedFocusArea: "",
      search: "",
      users: [],
      filteredUsers: [],
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      memberEvents: [],
      showUserActivities: false,
      userActivities: [],
      filteredUserActivities: [],
      userActivitiesPerPage: 5,
      pageNumber: 1,
      selectedMemberDetails: [],
      events: [],
      selectedMemberID: "",
      isExpanded: false,
      cardPersonRefs: []
    };
    //Create object for CommonServices class
    commonServiceManager = new commonServices(this.props.context, this.props.siteUrl);
    currentUserName = this.props.context.pageContext.user.displayName;
    this.onchange = this.onchange.bind(this);
    this._renderListAsync();
  }

  //Initializes the teams library and calling the methods to load the initial data  
  public _renderListAsync() {
    app.initialize();
    this.getMemberDetails();
    this.getChoicesFromList();
  }

  // Handle accordion toggle on tab - accessibility
  private handleAccordionToggle = (event: any, key: any) => {
    if (event.key === constants.stringEnter) {
      event.preventDefault();
      document.getElementById(`accordion-toggle-${key}`).click();
    }
  }

  // Set the expand collapse state - accessibility
  handleToggle = () => {
    const { isExpanded } = this.state;
    this.setState({ isExpanded: !isExpanded });
  };

  //This method will be called whenever there is an update to the component
  public componentDidUpdate(prevProps: Readonly<ChampionsCardsProps>, prevState: Readonly<ChampionsCardsState>, snapshot?: any): void {

    try {
      //Calling the methods to refresh the data in the champion cards when the component is re-rendered
      if (prevProps != this.props) {
        setTimeout(() => {
          this._renderListAsync();
        }, 500);
      }

      //Set the filteredUsers array based on the filter or search applied
      if ((prevState.selectedFocusArea != this.state.selectedFocusArea) ||
        (prevState.search != this.state.search) ||
        (prevState.users != this.state.users)) {
        let refLength: number;
        if (this.state.selectedFocusArea != constants.AllLabel && this.state.search == "") {
          //Filter the users based on the selected Focus Area and Searched value
          const filteredUsers = this.state.users.filter((user: any) => user.FocusArea?.toString().includes(this.state.selectedFocusArea));
          refLength = filteredUsers.length > this.state.loadCards ? this.state.loadCards : filteredUsers.length;
          //Set the filteredUsers array and cardPersonRefs array
          this.setState({
            filteredUsers: filteredUsers,
            cardPersonRefs: Array.from({ length: refLength }, () => React.createRef())
          });
        }
        else if (this.state.selectedFocusArea == constants.AllLabel && this.state.search == "") {
          refLength = this.state.users.length > this.state.loadCards ? this.state.loadCards : this.state.users.length;
          //Set the filteredUsers array and cardPersonRefs array
          this.setState({
            filteredUsers: this.state.users,
            cardPersonRefs: Array.from({ length: refLength }, () => React.createRef())
          });
        }
        else if (this.state.selectedFocusArea != constants.AllLabel && this.state.search != "") {
          //Filter the users based on the selected Focus Area and Searched value
          const filteredUsers = this.state.users.filter((user: any) =>
            user.FocusArea?.toString().includes(this.state.selectedFocusArea) &&
            ((user.FirstName &&
              user.FirstName.toLowerCase().includes(this.state.search.toLowerCase())) ||
              (user.LastName &&
                user.LastName.toLowerCase().includes(this.state.search.toLowerCase())) ||
              (user.Country &&
                user.Country.toLowerCase().includes(this.state.search.toLowerCase())) ||
              (user.FocusArea &&
                user.FocusArea?.toString().toLowerCase().includes(this.state.search.toLowerCase())) ||
              (user.Region &&
                user.Region.toLowerCase().includes(this.state.search.toLowerCase())) ||
              (user.Group && user.Group.toLowerCase().includes(this.state.search.toLowerCase()))));
          refLength = filteredUsers.length > this.state.loadCards ? this.state.loadCards : filteredUsers.length;
          //Set the filteredUsers array and cardPersonRefs array
          this.setState({
            filteredUsers: filteredUsers,
            cardPersonRefs: Array.from({ length: refLength }, () => React.createRef())
          });
        }
        else if (this.state.selectedFocusArea == constants.AllLabel && this.state.search != "") {
          //Filter the users based on the selected Focus Area and Searched value
          const filteredUsers = this.state.users.filter((user: any) =>
          ((user.FirstName &&
            user.FirstName.toLowerCase().includes(this.state.search.toLowerCase())) ||
            (user.LastName &&
              user.LastName.toLowerCase().includes(this.state.search.toLowerCase())) ||
            (user.Country &&
              user.Country.toLowerCase().includes(this.state.search.toLowerCase())) ||
            (user.FocusArea &&
              user.FocusArea?.toString().toLowerCase().includes(this.state.search.toLowerCase())) ||
            (user.Region &&
              user.Region.toLowerCase().includes(this.state.search.toLowerCase())) ||
            (user.Group && user.Group.toLowerCase().includes(this.state.search.toLowerCase()))));
          refLength = filteredUsers.length > this.state.loadCards ? this.state.loadCards : filteredUsers.length;
          //Set the filteredUsers array and cardPersonRefs array
          this.setState({
            filteredUsers: filteredUsers,
            cardPersonRefs: Array.from({ length: refLength }, () => React.createRef())
          });
        }
      }
      /**Update aria-expanded attribute in combobox for Accessibility in Android and 
        Add aria-label attribute to combobox label**/
      if (prevState.users !== this.state.users && this.state.users.length > 0) {
        const comboboxLabel = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listbox-label");
        //Add aria-label attribute to combobox label
        comboboxLabel.setAttribute("aria-label", LocaleStrings.SelectedFocusAreaLabel);

        //Update aria-expanded attribute in combobox for Accessibility in Android
        if (navigator.userAgent.match(/Android/i)) {

          //Outside Click event for Focus area combobox for Accessibility in Android
          document.addEventListener("click", this.onFocusAreaComboboxOutsideClick);

          //remove aria-expanded attribute from combobox input element
          const comboboxInput = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listbox-input");
          comboboxInput.removeAttribute("aria-expanded");

          //Update aria-expanded attribute for combobox expand/collapse button
          const comboboxButton = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listboxwrapper").querySelector("button");
          comboboxButton.setAttribute("aria-expanded", "false");

          //get focus area combobox list wrapper element to set focus and attributes
          const ulList: any = this.leaderboardFocusAreaComboboxRef?.current?.querySelector("#leaderboard-focus-area-listbox-list");
          ulList.setAttribute("tabindex", "0");
          comboboxButton.addEventListener("click", () => {
            setTimeout(() => {
              ulList.focus();
            }, 1000);
          });
        }
      }

      //Update the width of the card text block to break the text in multiple lines
      if (prevState.cardPersonRefs !== this.state.cardPersonRefs && this.state.cardPersonRefs.length > 0) {
        setTimeout(() => {
          for (let ele of this.state.cardPersonRefs) {
            const cardImgTextBlock = ele?.current?.shadowRoot?.querySelector(".person-root")
              ?.querySelector("mgt-flyout")?.querySelector(".details-wrapper")?.querySelector(".line1");
            cardImgTextBlock?.setAttribute("style", "width:195px;overflow-wrap: break-word;text-align: -webkit-auto");
          }
        }, 5000);
      }
    }
    catch (error) {
      console.error("CMP_ChampionsCards_componentDidUpdate_FailedToComponentUpdate \n", JSON.stringify(error));
    }
  }

  //Remove Document click event listener on Unmount of Component for Accessibility in Android
  public componentWillUnmount(): void {
    if (navigator.userAgent.match(/Android/i)) {
      document.removeEventListener("click", this.onFocusAreaComboboxOutsideClick);
    }
  }
  //Close Focus Area Combobox Callout on click of outside for Accessibility in Android
  public onFocusAreaComboboxOutsideClick = (evt: any) => {
    const isComboboxElement = document.getElementById("leaderboard-focus-area-listbox").contains(evt.target);
    if (!isComboboxElement) {
      this.mainComboboxRef.current.dismissMenu();
    }
  }

  //Get list of Champions from Member List and their Points from the Event Track Details list
  private async getMemberDetails() {

    let championsListArray: any = [];
    let eventTrackArray: any = [];
    let filteredMember: any = [];
    let allMemberEventsArray: any = [];

    //Get first batch of items from Event Track Details list
    let filterApprovedEvents = "Status eq 'Approved' or Status eq null or Status eq ''";
    let memberEventsArray = await commonServiceManager.getAllListItemsPagedWithFilter(constants.EventTrackDetailsList, filterApprovedEvents);
    if (memberEventsArray.results.length > 0) {
      allMemberEventsArray.push(...memberEventsArray.results);
      //Get next batch, if more items found in Event Track Details list
      while (memberEventsArray.hasNext) {
        memberEventsArray = await memberEventsArray.getNext();
        allMemberEventsArray.push(...memberEventsArray.results);
      }
      this.setState({
        memberEvents: allMemberEventsArray
      });
      eventTrackArray = allMemberEventsArray;
    }

    let filterQuery = "Status eq 'Approved'";
    let filter = "IsActive eq 1";
    await commonServiceManager.getItemsWithOnlyFilter(constants.MemberList, filterQuery)
      .then(async (approvedMembers) => {
        if (approvedMembers.length > 0) {
          await commonServiceManager.getItemsWithOnlyFilter(constants.EventsList, filter)
            .then(async (activeEvents) => {
              if (activeEvents.length > 0) {
                this.setState({
                  events: activeEvents
                });
                for (let i = 0; i < approvedMembers.length; i++) {
                  filteredMember = eventTrackArray.filter((user: any) => user.MemberId === approvedMembers[i].ID);
                  let eventpoints = _.groupBy(_.orderBy(filteredMember, ['Id'], ['asc']), "EventId");

                  let pointsCompleted: number = filteredMember.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue["Count"]; }, 0);
                  championsListArray.push({
                    Points: pointsCompleted,
                    EventPoints: eventpoints,
                    ID: approvedMembers[i].ID,
                    Title: approvedMembers[i].Title,
                    FirstName: approvedMembers[i].FirstName,
                    LastName: approvedMembers[i].LastName,
                    Region: approvedMembers[i].Region,
                    Country: approvedMembers[i].Country,
                    Role: approvedMembers[i].Role,
                    Status: approvedMembers[i].Status,
                    FocusArea: approvedMembers[i].FocusArea,
                    Group: approvedMembers[i].Group
                  });
                }
                //Sort by points                
                championsListArray.sort((a: any, b: any) => {
                  if (a.Points < b.Points) return 1;
                  if (a.Points > b.Points) return -1;
                });

                //Update ranks for the members
                championsListArray = championsListArray.map((currentValue: any, index: any) => {
                  currentValue.Rank = index + 1;
                  return currentValue;
                });

                this.setState({
                  users: championsListArray,
                  isLoaded: true,
                  selectedFocusArea: constants.AllLabel,
                  loadCards: constants.employeeCardLoadCount
                });
              }
            });
        }
      });
  }

  //Get dropdown choices for Region and Focus Area from Member List
  private async getChoicesFromList() {
    //Get choices for FocusArea dropdown from SharePoint list
    let focusAreas = await commonServiceManager.getChoicesFromListColumn(constants.MemberList, constants.FocusAreaColumn);
    this.setState({
      focusAreas: focusAreas
    });

  }

  //Adding the default choice "All"  for dropdown columns
  private options = (optionArray: any) => {
    let myOptions = [];
    myOptions.push({ key: constants.AllLabel, text: constants.AllLabel });
    optionArray.forEach((element: any) => {
      myOptions.push({ key: element, text: element });
    });
    return myOptions;
  }

  //Search the members based on the value entered in search box
  private onchange = (evt: any, value: string) => {
    if (value) {
      this.setState({
        search: value,
        filteredUsers: []
      });
    } else {
      this.setState({ search: "", filteredUsers: [] });
    }
  }

  //Setting state variable with the selected Focus Area
  private filterUsersByFocusArea = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({
      selectedFocusArea: option.key,
      filteredUsers: []
    });
  }

  //Method to execute the deep link API in teams
  public openTask = (selectedTask: string) => {
    app.initialize();
    app.openLink(selectedTask);
  }

  //Default image to show in case of any error in loading user profile image
  public addDefaultSrc(ev: any) {
    ev.target.src = require("../assets/images/noprofile.png");
  }

  /** On menu open add the attributes to fix the position issue in IOS and 
   Update aria-expanded attribute in combobox in Android for Accessibility **/
  private onMenuOpen = (listboxId: string) => {
    //Update aria-expanded attribute in combobox for Accessibility in Android
    if (navigator.userAgent.match(/Android/i)) {
      //remove aria-expanded attribute from combobox input element
      const comboboxInput = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listbox-input");
      comboboxInput.removeAttribute("aria-expanded");

      //Update aria-expanded attribute for combobox expand/collapse button
      const comboboxButton = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listboxwrapper").querySelector("button");
      comboboxButton.setAttribute("aria-expanded", "true");
    }

    //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
    if (navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) {
      const listBoxElement: any = document.getElementById(listboxId + "-list")?.children;
      if (listBoxElement?.length > 0) {
        for (let i = 0; i < listBoxElement?.length; i++) {
          const buttonId = `${listboxId}-list${i}`;
          const buttonElement: any = document.getElementById(buttonId);
          const ariaLabel = `${buttonElement.innerText} ${i + 1} of ${listBoxElement.length}`;
          buttonElement?.setAttribute("aria-label", ariaLabel);
        }
      }
    }

  }

  //Main render method
  public render() {
    const isDarkOrContrastTheme = this.props.currentThemeName === constants.themeDarkMode || this.props.currentThemeName === constants.themeContrastMode;
    return (
      <React.Fragment>
        {this.state.isLoaded && this.state.users.length === 0 && (
          <div className={`m-4 card${isDarkOrContrastTheme ? " no-results--DarkContrast" : ""}`}>
            <b
              className='card-title p-4 text-center'
              aria-live="polite" role="alert"
            >
              {LocaleStrings.RecordsNotFound}</b>
          </div>
        )}
        {this.state.users.length > 0 && (
          <>
            <div className={`championsFilterArea${isDarkOrContrastTheme ? " championsFilterAreaDarkContrast" : ""}`}>
              <Row xl={2} lg={2} md={2} sm={1} xs={1}>
                <Col xl={12} lg={12} md={12} sm={12} xs={12}>
                  <h1 tabIndex={0} role="heading"><div className="topChampionsLabel">
                    {LocaleStrings.TopChampionsLabel}
                  </div></h1>
                </Col>
                <Col xl={5} lg={6} md={12} sm={12} xs={12}>
                  <div className="championFocusAreaComboboxArea" ref={this.leaderboardFocusAreaComboboxRef}>
                    <ComboBox
                      label={LocaleStrings.FocusAreaLabel}
                      selectedKey={this.state.selectedFocusArea}
                      options={this.options(this.state.focusAreas)}
                      onChange={this.filterUsersByFocusArea.bind(this)}
                      className="championFocusAreaCombobox"
                      ariaLabel={LocaleStrings.FocusAreaLabel}
                      useComboBoxAsMenuWidth={true}
                      calloutProps={{
                        className: "championFocusAreaCallout", directionalHintFixed: true, doNotLayer: true,
                        preventDismissOnEvent: () => {
                          //Prevent callout closing in Android on very first time opening it for Accessibility
                          if (navigator.userAgent.match(/Android/i)) {
                            return true;
                          }
                          else {
                            return false;
                          }
                        }
                      }}
                      allowFreeInput={true}
                      persistMenu={true}
                      id="leaderboard-focus-area-listbox"
                      onMenuOpen={() => this.onMenuOpen("leaderboard-focus-area-listbox")}
                      onMenuDismissed={() => {
                        //Update aria-expanded attribute in combobox for Accessibility in Android
                        if (navigator.userAgent.match(/Android/i)) {
                          //remove aria-expanded attribute from combobox input element
                          const comboboxInput = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listbox-input");
                          comboboxInput.removeAttribute("aria-expanded");

                          //Update aria-expanded attribute for combobox expand/collapse button
                          const comboboxButton = this.leaderboardFocusAreaComboboxRef.current.querySelector("#leaderboard-focus-area-listboxwrapper").querySelector("button");
                          comboboxButton.setAttribute("aria-expanded", "false");
                        }
                      }}
                      componentRef={this.mainComboboxRef}
                    />
                  </div>
                </Col>
                <Col xl={5} lg={6} md={12} sm={12} xs={12}>
                  <div className="championSearchboxArea">
                    <SearchBox
                      className="championSearchbox"
                      placeholder={LocaleStrings.SearchLabel}
                      onChange={this.onchange}
                    />
                  </div>
                </Col>
              </Row>
            </div>
            {this.state.filteredUsers.length > 0 && (navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
              <div aria-live="polite" role="alert" className={`records-count-label${isDarkOrContrastTheme ? " records-count-labelDarkContrast" : ""}`}>
                {this.state.filteredUsers.length} {LocaleStrings.championRecordsFoundLabel}
              </div>
            }
            <div className='gtc-cards'>
              <Row xl={3} lg={2} md={1} sm={1} xs={1}>
                {(this.state.filteredUsers.filter((_user: any, idx: number) => idx < this.state.loadCards))
                  .map((member: any, ind: number) => {
                    return (
                      this.state.isLoaded && (
                        <Col xl={4} lg={6} md={12} sm={12} xs={12}>
                          <div className="cards">
                            <div className="card-img-text-block">
                              <div>
                                <Person
                                  personQuery={member.Title}
                                  view={4}
                                  personCardInteraction={1}
                                  avatarSize="large"
                                  ref={this.state.cardPersonRefs[ind]}
                                />
                                <div
                                  className={`rank-points-block${this.props.loggedinUserEmail === member.Title ? " highlight-data" : ""}`}
                                >
                                  <span className="card-rank" title={`Rank ${member.Rank}`}>{LocaleStrings.RankLabel} <span className="card-rank-value">#{member.Rank}</span></span>
                                  <span className="card-points" title={`${member.Points ? member.Points : ""} ${LocaleStrings.PointsLabel}`}>
                                    {member.Points}
                                    <Icon iconName="FavoriteStarFill" className="card-points-star" />
                                  </span>
                                </div>
                              </div>
                            </div>

                            <div className={`card-icon-link-area${this.props.loggedinUserEmail === member.Title ? " align-link-end" : ""}`}>
                              {this.props.loggedinUserEmail !== member.Title &&
                                <div
                                  className="request-to-call-link"
                                  title={LocaleStrings.RequestToCallLabel}
                                  onClick={() => this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                    currentUserName + " / " + member.FirstName + " " + member.LastName + " " + LocaleStrings.MeetupSubject +
                                    "&content=" + LocaleStrings.MeetupBody + "&attendees=" + member.Title)}
                                  onKeyDown={(evt: any) => {
                                    if (evt.key === constants.stringEnter) this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                      currentUserName + " / " + member.FirstName + " " + member.LastName + " " + LocaleStrings.MeetupSubject +
                                      "&content=" + LocaleStrings.MeetupBody + "&attendees=" + member.Title)
                                  }}
                                  tabIndex={0}
                                >
                                  {LocaleStrings.RequestToCallLabel}
                                </div>
                              }
                              <div
                                className="view-activities-link"
                                title={LocaleStrings.ViewActivitiesLabel}
                                onClick={() => {
                                  this.setState({
                                    showUserActivities: true,
                                    selectedMemberDetails: member,
                                    selectedMemberID: member.ID
                                  })
                                }}
                                onKeyDown={(evt: any) => {
                                  if (evt.key === constants.stringEnter) this.setState({
                                    showUserActivities: true,
                                    selectedMemberDetails: member,
                                    selectedMemberID: member.ID
                                  })
                                }}
                                tabIndex={0}
                              >
                                {LocaleStrings.ViewActivitiesLabel}
                              </div>

                            </div>
                          </div>
                        </Col>
                      )
                    );
                  })}
                {
                  (this.state.loadCards < this.state.filteredUsers.length && this.state.isLoaded) &&
                  <Col xl={4} lg={6} md={12} sm={12} xs={12}>
                    <div className={`cards-show-label${isDarkOrContrastTheme ? " cards-show-labelDarkContrast" : ""}`}>
                      <span
                        onClick={() => {
                          const cardCount = this.state.loadCards + constants.employeeCardLoadCount;
                          const refLength = this.state.filteredUsers.length > cardCount ? cardCount : this.state.filteredUsers.length;
                          this.setState({ loadCards: cardCount, cardPersonRefs: Array.from({ length: refLength }, () => React.createRef()) });
                        }}
                        title={LocaleStrings.ShowMoreLabel}
                        tabIndex={0}
                        onKeyDown={(evt: any) => {
                          if (evt.key === constants.stringEnter) {
                            const cardCount = this.state.loadCards + constants.employeeCardLoadCount;
                            const refLength = this.state.filteredUsers.length > cardCount ? cardCount : this.state.filteredUsers.length;
                            this.setState({ loadCards: cardCount, cardPersonRefs: Array.from({ length: refLength }, () => React.createRef()) });
                          }
                        }}
                        className='show-more-text-img-wrapper'
                      >
                        <span aria-hidden="true">{LocaleStrings.ShowMoreLabel}</span>
                        <img src={require("../assets/CMPImages/ShowMoreIcon.svg")} alt="" className="showMoreIcon" aria-hidden={true} />
                      </span>
                    </div>
                  </Col>
                }
              </Row>
              {this.state.filteredUsers.length > 0 && !(navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
                <span aria-label={`${this.state.filteredUsers.length} ${LocaleStrings.championRecordsFoundLabel}`} aria-live="polite" role="alert" />
              }
              {this.state.isLoaded && this.state.filteredUsers.length === 0 && (
                <div className={`m-4 card${isDarkOrContrastTheme ? " no-results--DarkContrast" : ""}`}>
                  <b
                    className='card-title p-4 text-center'
                    aria-live="polite" role="alert">
                    {LocaleStrings.RecordsNotFound}</b>
                </div>
              )}
              {this.state.showUserActivities &&
                <Dialog
                  hidden={!this.state.showUserActivities}
                  onDismiss={() => this.setState({ showUserActivities: false })}
                  modalProps={{
                    isBlocking: false,
                    className: `showActivitiesPopup${isDarkOrContrastTheme ? " " + this.props.currentThemeName + "Popup" : ""}`
                  }}
                  dialogContentProps={{ showCloseButton: false }}
                >
                  <div className="chrome-close-icon-area">
                    <Icon
                      iconName="ChromeClose"
                      className="chrome-close-icon"
                      onClick={() => this.setState({ showUserActivities: false })}
                      tabIndex={0}
                      onKeyDown={(evt: any) => { if (evt.key === constants.stringEnter) this.setState({ showUserActivities: false }) }}
                    />
                  </div>
                  <ChampionEvents
                    context={this.props.context}
                    filteredAllEvents={this.state.memberEvents}
                    selectedMemberDetails={this.state.selectedMemberDetails}
                    parentComponent={constants.ChampionsCardsLabel}
                    selectedMemberID={this.state.selectedMemberID}
                    loggedinUserEmail={this.props.loggedinUserEmail}
                  />
                </Dialog>
              }</div>

            <div className="paddingTop">
              {this.props.type && (
                <React.Fragment>
                  <span className="topChampionsLabel">
                    {LocaleStrings.TopChampionsLabel}
                    : <b>{LocaleStrings.MyRankLabel}</b>
                  </span>
                  <div className="table-content">
                    {this.state.isLoaded && (
                      <Accordion>
                        {this.state.users
                          .slice(0, 3)
                          .map((rankedMember: any, ind: number) => {
                            return (
                              <Card className={`topChampCards${isDarkOrContrastTheme ? " topChampCards--DarkContrast" : ""}`} key={rankedMember.ID}>
                                <Accordion.Toggle
                                  as={Card.Header}
                                  eventKey={rankedMember.ID}
                                  tabIndex={0}
                                  onKeyDown={(event: any) => this.handleAccordionToggle(event, '2')}
                                  id="accordion-toggle-2"
                                  role="button"
                                  aria-expanded={this.state.isExpanded}
                                  onClick={this.handleToggle}
                                >
                                  <div className="gttc-row-left">
                                    <div className='gttc-img'>
                                      <Person
                                        personQuery={rankedMember.Title}
                                        view={3}
                                        personCardInteraction={1}
                                        className="accordion-person-card"
                                      />
                                    </div>
                                  </div>
                                  <div className="gttc-row-right">
                                    <div className="gttc-star">
                                      <Icon
                                        iconName="FavoriteStarFill"
                                        id="points2"
                                      />
                                      <span className="points">{rankedMember.Points}</span>
                                    </div>
                                    <div className="vline"></div>
                                    <div className="gttc-rank">
                                      {LocaleStrings.RankLabel} <b>{rankedMember.Rank}</b>
                                    </div>
                                  </div>
                                </Accordion.Toggle>
                                <Accordion.Collapse eventKey={rankedMember.ID} className={`${isDarkOrContrastTheme ? " accordion-collapse--DarkContrast" : ""}`}>
                                  <Card.Body>
                                    <Table>
                                      {Object.keys(rankedMember.EventPoints).length != 0 &&
                                        <tr>
                                          <th>{LocaleStrings.EventTypeLabel}</th>
                                          <th className="countHeader">{LocaleStrings.CountLabel}</th>
                                        </tr>
                                      }
                                      {Object.keys(rankedMember.EventPoints).map((e, i) => {
                                        return (
                                          e !== "0" && (
                                            <tr>
                                              <td className="eventTypeCol">
                                                {
                                                  this.state.events.find((ev) => e !== "0" && ev.ID.toString() === e)
                                                  && this.state.events.find((ev) => e !== "0" && ev.ID.toString() === e).Title
                                                }
                                              </td>
                                              <td className="gttc-tap-data">
                                                {
                                                  rankedMember.EventPoints[e].map((x: any) => x.Count / x.Count).length
                                                }
                                              </td>
                                            </tr>
                                          )
                                        );
                                      })}
                                    </Table>
                                  </Card.Body>
                                </Accordion.Collapse>
                              </Card>
                            );
                          })}
                      </Accordion>
                    )}
                  </div>
                </React.Fragment>
              )}
            </div>
          </>
        )}
      </React.Fragment>
    );
  }
}
