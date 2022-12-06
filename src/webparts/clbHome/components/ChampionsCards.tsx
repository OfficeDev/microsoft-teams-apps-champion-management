import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as microsoftTeams from '@microsoft/teams-js';
import { Component } from 'react';
import Accordion from 'react-bootstrap/Accordion';
import BootstrapTable from 'react-bootstrap-table-next';
import Card from 'react-bootstrap/Card';
import Col from 'react-bootstrap/esm/Col';
import Row from 'react-bootstrap/Row';
import Table from 'react-bootstrap/Table';
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { Dialog } from '@fluentui/react/lib/Dialog';
import { Icon, initializeIcons } from 'office-ui-fabric-react';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import moment from 'moment';
import * as constants from '../constants/strings';
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import siteconfig from '../config/siteconfig.json';
import '../scss/Championview.scss';
import '../scss/Champions.scss';
import * as _ from "lodash";
import commonServices from '../Common/CommonServices';

initializeIcons();

let commonServiceManager: commonServices;
let currentUserName: string;

interface ChampionsCardsProps {
  type: string;
  events?: any;
  context: WebPartContext;
  siteUrl: string;
  callBack?: Function;
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
  selectedMemberDetails: Array<any>;
  events: Array<any>;
}

export default class ChampionsCards extends Component<
  ChampionsCardsProps,
  ChampionsCardsState
> {
  constructor(props: any) {
    super(props);
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
      events: []
    };
    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
    currentUserName = this.props.context.pageContext.user.displayName;
    this.onchange = this.onchange.bind(this);
    this._renderListAsync();
  }

  //Initializes the teams library and calling the methods to load the initial data  
  public _renderListAsync() {
    microsoftTeams.initialize();
    this.getMemberDetails();
    this.getChoicesFromList();
  }

  //This method will be called whenever there is an update to the component
  public componentDidUpdate(prevProps: Readonly<ChampionsCardsProps>, prevState: Readonly<ChampionsCardsState>, snapshot?: any): void {
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

      if (this.state.selectedFocusArea != constants.AllLabel && this.state.search == "") {
        this.setState({
          filteredUsers: this.state.users.filter((user) => user.FocusArea?.toString().includes(this.state.selectedFocusArea))
        });
      } else if (this.state.selectedFocusArea == constants.AllLabel && this.state.search == "") {
        this.setState({
          filteredUsers: this.state.users
        });
      } else if (this.state.selectedFocusArea != constants.AllLabel && this.state.search != "") {
        this.setState({
          filteredUsers: this.state.users.filter((user) =>
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
              (user.Group && user.Group.toLowerCase().includes(this.state.search.toLowerCase()))))
        });
      } else if (this.state.selectedFocusArea == constants.AllLabel && this.state.search != "") {
        this.setState({
          filteredUsers: this.state.users.filter((user) =>
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
            (user.Group && user.Group.toLowerCase().includes(this.state.search.toLowerCase()))))
        });
      }
    }

    if (prevState.userActivities.length !== this.state.userActivities.length || prevState.pageNumber !== this.state.pageNumber) {
      this.updatefilteredUserActivities();
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
                  filteredMember = eventTrackArray.filter(user => user.MemberId === approvedMembers[i].ID);
                  let eventpoints = _.groupBy(_.orderBy(filteredMember, ['Id'], ['asc']), "EventId");

                  let pointsCompleted: number = filteredMember.reduce((previousValue, currentValue) => { return previousValue + currentValue["Count"]; }, 0);
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
                championsListArray.sort((a, b) => {
                  if (a.Points < b.Points) return 1;
                  if (a.Points > b.Points) return -1;
                });

                //Update ranks for the members
                championsListArray = championsListArray.map((currentValue, index) => {
                  currentValue.Rank = index + 1;
                  return currentValue;
                });

                this.setState({
                  users: championsListArray,
                  isLoaded: true,
                  selectedFocusArea: constants.AllLabel,
                  loadCards: constants.employeeCardLoadCount,
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

  //Get the selected member's activities and their user data to show it in the modal popup
  private getMemberActivities(selectedMember: any, rank: any) {
    let memberActivitesArray: any = [];
    let memberDetails: any = [];

    //Filtering the selected member's data from the array of all records from Event Track Details list
    let selectedMemberEvents = this.state.memberEvents.filter(user => user.MemberId === selectedMember.ID);

    //Creating an array to store the required data for Activities table in the popup screen
    selectedMemberEvents.forEach((event) => {
      memberActivitesArray.push({
        DateofEvent: moment(event["DateofEvent"]).format("MMMM Do, YYYY"),
        Type: event["EventName"],
        Points: event["Count"]
      });
    });

    //Creating an array to store the user data of the selected member to display it in the popup
    memberDetails.push({
      Points: selectedMember.Points,
      ID: selectedMember.ID,
      Title: selectedMember.Title,
      FirstName: selectedMember.FirstName,
      LastName: selectedMember.LastName,
      Rank: rank
    });
    this.setState({
      userActivities: memberActivitesArray,
      selectedMemberDetails: memberDetails
    });
  }

  //Filtering the records based on page size for each page from total activities of the member
  private updatefilteredUserActivities = () => {
    const filteredData = this.state.userActivities.filter((activity, idx) => {
      return (idx >= (this.state.userActivitiesPerPage * this.state.pageNumber - this.state.userActivitiesPerPage) && idx < (this.state.pageNumber * this.state.userActivitiesPerPage));
    });
    this.setState({ filteredUserActivities: filteredData });
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
        search: value
      });
    } else {
      this.setState({ search: "" });
    }
  }

  //Setting state variable with the selected Focus Area
  private filterUsersByFocusArea = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({
      selectedFocusArea: option.key
    });
  }

  //Method to execute the deep link API in teams
  public openTask = (selectedTask: string) => {
    microsoftTeams.initialize();
    microsoftTeams.executeDeepLink(selectedTask);
  }

  //Default image to show in case of any error in loading user profile image
  public addDefaultSrc(ev: any) {
    ev.target.src = require("../assets/images/noprofile.png");
  }

  //Main render method
  public render() {
    const starStyles = {
      color: "#f3ca3e"
    };
    const activitiesTableHeader = [
      {
        dataField: 'DateofEvent',
        text: 'Date of Events',
        headerTitle: true,
        title: true,
      },
      {
        dataField: 'Type',
        text: 'Type',
        headerTitle: true,
        title: true,
      },
      {
        dataField: 'Points',
        text: 'Points',
        headerTitle: true,
        title: true,
      }
    ];
    return (
      <React.Fragment>
        {this.state.isLoaded && this.state.users.length === 0 && (
          <div className="m-4 card">
            <b className="card-title p-4 text-center">{LocaleStrings.RecordsNotFound}</b>
          </div>
        )}
        {this.state.users.length > 0 && (
          <>
            <div className="championsFilterArea">
              <Row xl={2} lg={2} md={2} sm={1} xs={1}>
                <Col xl={12} lg={12} md={12} sm={12} xs={12}>
                  <div className="topChampionsLabel">
                    {LocaleStrings.TopChampionsLabel}
                  </div>
                </Col>
                <Col xl={5} lg={6} md={12} sm={12} xs={12}>
                  <div className="championFocusAreaComboboxArea">
                    <ComboBox
                      label={LocaleStrings.FocusAreaLabel}
                      selectedKey={this.state.selectedFocusArea}
                      options={this.options(this.state.focusAreas)}
                      onChange={this.filterUsersByFocusArea.bind(this)}
                      className="championFocusAreaCombobox"
                      calloutProps={{ className: "championFocusAreaCallout" }}
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
            <div className="gtc-cards">
              <Row xl={3} lg={2} md={1} sm={1} xs={1}>
                {(this.state.filteredUsers.filter((user, idx) => idx < this.state.loadCards))
                  .map((member: any, ind = 0) => {
                    return (
                      this.state.isLoaded && (
                        <Col xl={4} lg={6} md={12} sm={12} xs={12}>
                          <div className="cards">
                            <div className="card-img-text-block">
                              <div>
                                <img
                                  src={
                                    "/_layouts/15/userphoto.aspx?size=L&username=" +
                                    member.Title
                                  }
                                  className="profile-img"
                                  onError={this.addDefaultSrc}
                                  alt={member.FirstName}
                                  title={member.FirstName}
                                />
                              </div>
                              <div>
                                <div className="gtc-name2" title={member.FirstName}>
                                  {member.FirstName}{' '}{member.LastName}
                                </div>
                                <div className="rank-points-block">
                                  <span className="card-rank" title={`Rank ${ind + 1}`}>Rank <span className="card-rank-value">#{member.Rank}</span></span>
                                  <span className="card-points" title={`${member.Points ? member.Points : ""} Points`}>
                                    {member.Points}
                                    <Icon iconName="FavoriteStarFill" className="card-points-star" />
                                  </span>
                                </div>
                              </div>
                            </div>

                            <div className="card-icon-link-area">
                              <div className="card-icon-area">
                                <img
                                  src={require("../assets/CMPImages/EmployeeChatIcon.svg")}
                                  alt="Employee Chat Icon"
                                  className="card-icon"
                                  title={LocaleStrings.ChatIconLabel}
                                  onClick={() => this.openTask(`https://teams.microsoft.com/l/chat/0/0?users=${member.Title}`)}
                                />
                                <img
                                  src={require("../assets/CMPImages/CallRequestIcon.svg")}
                                  alt="Call Request Icon"
                                  className="card-icon"
                                  title={LocaleStrings.RequestToCallLabel}
                                  onClick={() => this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                    currentUserName + " / " + member.FirstName + " " + member.LastName + " " + LocaleStrings.MeetupSubject +
                                    "&content=" + LocaleStrings.MeetupBody + "&attendees=" + member.Title)}
                                />
                                <a href={`mailto:${member.Title}`}>
                                  <img
                                    src={require("../assets/CMPImages/EmployeeMailIcon.svg")}
                                    alt="Employee Mail Icon"
                                    className="card-icon"
                                    title={LocaleStrings.EmailIconLabel}
                                  />
                                </a>
                              </div>
                              <div
                                className="card-link-area"
                                title={LocaleStrings.ViewActivitiesLabel}
                                onClick={() => { this.setState({ showUserActivities: true }); this.getMemberActivities(member, ind + 1); }}
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
                    <div
                      onClick={() => this.setState({ loadCards: this.state.loadCards + constants.employeeCardLoadCount })}
                      className="cards-show-label"
                      title={LocaleStrings.ShowMoreLabel}
                    > {LocaleStrings.ShowMoreLabel} <img src={require("../assets/CMPImages/ShowMoreIcon.svg")} alt="" className="showMoreIcon" /></div>
                  </Col>
                }
              </Row>
              {this.state.isLoaded && this.state.filteredUsers.length === 0 && (
                <div className="m-4 card">
                  <b className="card-title p-4 text-center">{LocaleStrings.RecordsNotFound}</b>
                </div>
              )}
              {this.state.showUserActivities &&
                <Dialog
                  hidden={!this.state.showUserActivities}
                  onDismiss={() => this.setState({ showUserActivities: false, pageNumber: 1 })}
                  modalProps={{ isBlocking: false }}
                  className="showActivitiesPopup"
                >
                  <div className="chrome-close-icon-area">
                    <Icon
                      iconName="ChromeClose"
                      className="chrome-close-icon"
                      onClick={() => this.setState({ showUserActivities: false, pageNumber: 1 })}
                      tabIndex={0}
                    />
                  </div>
                  <div className="showActivitiesPopupBody">
                    <Row xl={2} lg={2} md={1} sm={1} xs={1}>
                      <Col xl={4} lg={4} md={12} sm={12} xs={12}>
                        <div className="showActivitiesImage-IconArea">
                          <img
                            src={
                              "/_layouts/15/userphoto.aspx?size=L&username=" +
                              this.state.selectedMemberDetails[0].Title
                            }
                            className="showActivities-profile-img"
                            onError={this.addDefaultSrc}
                            alt={this.state.selectedMemberDetails[0].FirstName}
                            title={this.state.selectedMemberDetails[0].FirstName}
                          />
                          <div className="showActivities-profile-name">
                            {this.state.selectedMemberDetails[0].FirstName}{' '}{this.state.selectedMemberDetails[0].LastName}
                          </div>
                          <div className="showActivities-rank-points-block">
                            <span className="showActivities-rank" title={`Rank 1`}>Rank <span className="showActivities-rank-value"># {this.state.selectedMemberDetails[0].Rank}</span></span>
                            <span className="showActivities-points" title={`#Points`}>
                              {this.state.selectedMemberDetails[0].Points}
                              <Icon iconName="FavoriteStarFill" className="showActivities-points-star" />
                            </span>
                          </div>
                          <div className="showActivities-icon-area">
                            <img
                              src={require("../assets/CMPImages/EmployeeChatIcon.svg")}
                              alt="Employee Chat Icon"
                              className="showActivities-icon"
                              title={LocaleStrings.ChatIconLabel}
                              onClick={() => this.openTask(`https://teams.microsoft.com/l/chat/0/0?users=${this.state.selectedMemberDetails[0].Title}`)}
                            />
                            <img
                              src={require("../assets/CMPImages/CallRequestIcon.svg")}
                              alt="Call Request Icon"
                              className="showActivities-icon"
                              title={LocaleStrings.RequestToCallLabel}
                              onClick={() => this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                currentUserName + " / " + this.state.selectedMemberDetails[0].FirstName + " " + this.state.selectedMemberDetails[0].LastName + " " + LocaleStrings.MeetupSubject +
                                "&content=" + LocaleStrings.MeetupBody + "&attendees=" + this.state.selectedMemberDetails[0].Title)}
                            />
                            <a href={`mailto:${this.state.selectedMemberDetails[0].Title}`}>
                              <img
                                src={require("../assets/CMPImages/EmployeeMailIcon.svg")}
                                alt="Employee Mail Icon"
                                className="showActivities-icon"
                                title={LocaleStrings.EmailIconLabel}
                              />
                            </a>
                          </div>
                        </div>
                      </Col>
                      <Col xl={8} lg={8} md={12} sm={12} xs={12}>
                        <div className="showActivities-grid-area">
                          <div className="activities-grid-heading">{LocaleStrings.ActivitiesLabel}</div>
                          <BootstrapTable
                            bootstrap4
                            keyField={'dateOfEvents'}
                            data={this.state.filteredUserActivities}
                            columns={activitiesTableHeader}
                            table-responsive={true}
                            noDataIndication={() => (<div className='activities-noRecordsFound'>{LocaleStrings.NoActivitiesinGridLabel}</div>)}
                          />
                          {this.state.filteredUserActivities.length > 0 &&
                            <div className="pagination-area">
                              <span>
                                {this.state.pageNumber} of {Math.ceil(this.state.userActivities.length / this.state.userActivitiesPerPage)}

                                <Icon
                                  iconName="ChevronLeft"
                                  className="Chevron-Icon"
                                  onClick={this.state.pageNumber > 1 ? () => { this.setState({ pageNumber: this.state.pageNumber - 1 }); } : null}
                                />

                                <Icon
                                  iconName="ChevronRight"
                                  className="Chevron-Icon"
                                  onClick={this.state.pageNumber < Math.ceil(this.state.userActivities.length / this.state.userActivitiesPerPage) ? () => { this.setState({ pageNumber: this.state.pageNumber + 1 }); } : null}
                                />
                              </span>
                            </div>
                          }
                        </div>
                      </Col>
                    </Row>
                  </div>
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
                              <Card className="topChampCards">
                                <Accordion.Toggle
                                  as={Card.Header}
                                  eventKey={rankedMember.ID}
                                >
                                  <div className="gttc-row-left">
                                    <span>
                                      <img src={"/_layouts/15/userphoto.aspx?size=M&username=" + rankedMember.Title}
                                        className="gttc-img"
                                        onError={this.addDefaultSrc}
                                        alt={rankedMember.FirstName}
                                      />
                                      <div className="gttc-img-name">
                                        {rankedMember.FirstName}
                                      </div>
                                    </span>
                                  </div>
                                  <div className="gttc-row-right">
                                    <div className="gttc-star">
                                      <Icon
                                        iconName="FavoriteStarFill"
                                        id="points2"
                                        style={starStyles}
                                      />
                                      <span className="points">{rankedMember.Points}</span>
                                    </div>
                                    <div className="vline"></div>
                                    <div className="gttc-rank">
                                      {LocaleStrings.RankLabel} <b>{ind + 1}</b>
                                    </div>
                                  </div>
                                </Accordion.Toggle>
                                <Accordion.Collapse eventKey={rankedMember.ID}>
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
                                                  rankedMember.EventPoints[e].map((x) => x.Count / x.Count).length
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
