import { Icon } from '@fluentui/react/lib/Icon';
import {
  DatePicker,
  DayOfWeek, IDatePickerStrings
} from "@fluentui/react/node_modules/office-ui-fabric-react/lib/DatePicker";
import { DataGrid } from "@material-ui/data-grid";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import cx from "classnames";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import _ from "lodash";
import * as moment from "moment";
import { DefaultButton } from "office-ui-fabric-react";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import React, { Component } from "react";
import Accordion from "react-bootstrap/Accordion";
import Card from "react-bootstrap/Card";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import commonServices from '../Common/CommonServices';
import Sidebar from "../components/Sidebar";
import siteconfig from "../config/siteconfig.json";
import "../scss/Championview.scss";

const columns = [
  {
    field: "DateOfEvent", type: 'date', sortable: false,
    headerName: LocaleStrings.DateofEventGridLabel, width: 200
  },
  { field: "type", headerName: LocaleStrings.EventTypeGridLabel, width: 150 },
  { field: "Count", type: 'number', headerName: LocaleStrings.CMPSideBarPointsLabel, width: 150 },
];
const DayPickerStrings: IDatePickerStrings = {
  months: [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ],

  shortMonths: [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ],

  days: [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ],

  shortDays: ["S", "M", "T", "W", "T", "F", "S"],

  goToToday: "Go to today",
  prevMonthAriaLabel: "Go to previous month",
  nextMonthAriaLabel: "Go to next month",
  prevYearAriaLabel: "Go to previous year",
  nextYearAriaLabel: "Go to next year",
  closeButtonAriaLabel: "Close date picker",
};

const firstDayOfWeek = DayOfWeek.Sunday;
export interface ChampionViewProps {
  context: WebPartContext;
  onClickCancel: () => void;
  showSidebar?: boolean;
  callBack?: Function;
  siteUrl: string;
}
export interface lookUp {
  value: string;
  display: string;
}
export interface ChampionViewState {
  siteUrl: string;
  type: string;
  teams: Array<lookUp>;
  selectedTeam: string;
  validationError: string;
  eventid: number;
  memberid: number;
  points: number;
  DateOfEvent: Date;
  collection: Array<ChampList>;
  collectionNew: Array<ChampList>;
  edetails: Array<string>;
  edetailsIds: Array<EventList>;
  eFlag: boolean;
  optionvalues: Array<string>;
  selectedkey: number;
  isShow: boolean;
  cb: boolean;
  Clb: boolean;
  newMemberId: number;
  sitename: string;
  inclusionpath: string;
  loading: boolean;
  membersInfo: Array<any>;
  showValidationError: boolean;
  eventUniqueID: number;
}
export interface ChampList {
  id: number;
  type: string;
  eventid: number;
  memberid: number;
  Count: number;
  DateOfEvent: any;
  MemberName: string;
  EventName: string;
}
export interface EventList {
  Title: string;
  Id: number;
  Ecount: number;
}
export default class ChampionvView extends Component<
  ChampionViewProps,
  ChampionViewState
> {
  constructor(props: any) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });
    this.getTrackDetailsData = this.getTrackDetailsData.bind(this);
    this.onChange = this.onChange.bind(this);
    this.getListData = this.getListData.bind(this);
    this.setPoints = this.setPoints.bind(this);
    this.createorupdateItem = this.createorupdateItem.bind(this);
    this.options = this.options.bind(this);
    this.removeDevice = this.removeDevice.bind(this);
    this.getMemberId = this.getMemberId.bind(this);
    this.state = {
      siteUrl: this.props.siteUrl,
      type: "",
      teams: [],
      selectedTeam: "",
      validationError: "",
      eventid: 0,
      memberid: 0,
      points: 1,
      DateOfEvent: new Date(),
      collection: [],
      collectionNew: [],
      edetails: [],
      edetailsIds: [],
      eFlag: false,
      optionvalues: [],
      selectedkey: 0,
      isShow: false,
      cb: false,
      Clb: false,
      newMemberId: 0,
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      loading: true,
      membersInfo: [],
      showValidationError: false,
      eventUniqueID: 0
    };
  }

  public onChange(d: any) {
    this.setState({ DateOfEvent: d });
  }

  //When a new event is added modify the collection to show in the grid
  public addDevice(data: ChampList, saved: any) {
    if (saved === "false") {
      if ((data.type == "" || data.type == "Select Event Type")) {
        this.setState({ showValidationError: true, validationError: LocaleStrings.EventTypeValidationMessage });
      }
      else if ((data.Count > 5 || data.Count < 1)) {
        this.setState({ showValidationError: true, validationError: LocaleStrings.CountValidationMessage });
      }
      else {
        this.setState({ collectionNew: [], showValidationError: false, eventUniqueID: data.id });
        const newBag = this.state.collectionNew.concat(data);

        this.setState({
          collectionNew: newBag,
          points: data.Count,
        });
        this.setState({ selectedkey: 0 });
      }

    } else {
      const newBag = this.state.collection.concat(data);
      this.setState({
        collection: newBag,
      });
    }
  }
  public options() {
    let optionArray = [];
    let optionArrayIds = [];
    if (this.state.edetails.length == 0)
      this.props.context.spHttpClient
        .get(

          "/" +
          this.state.inclusionpath +
          "/" +
          this.state.sitename +
          "/_api/web/lists/GetByTitle('Events List')/Items",
          SPHttpClient.configurations.v1
        )
        .then(async (response: SPHttpClientResponse) => {
          if (response.status === 200) {
            await response.json().then((responseJSON: any) => {
              let i = 0;
              while (i < responseJSON.value.length) {
                if (
                  responseJSON.value[i] &&
                  responseJSON.value[i].hasOwnProperty("Title") &&
                  responseJSON.value[i].IsActive
                ) {
                  optionArray.push(responseJSON.value[i].Title);
                  optionArrayIds.push({
                    Title: responseJSON.value[i].Title,
                    Id: responseJSON.value[i].Id,
                    Ecount: responseJSON.value[i].Points,
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

    let myOptions = [];
    myOptions.push({ key: "Select Event Type", text: "Select Event Type" });
    this.state.edetails.forEach((element: any) => {
      myOptions.push({ key: element, text: element });
    });
    return myOptions;
  }

  //Remove the record from the grid when clicked on delete icon
  public removeDevice(points: number, recordID: number) {
    this.setState((prevState) => ({
      collectionNew: prevState.collectionNew.filter(
        (data) => data.id !== recordID,
        points
      ),
    }));
  }

  public createorupdateItem() {
    this.setState({ isShow: true });
    let commonServiceManager: commonServices = new commonServices(this.props.context, this.props.siteUrl);
    let promiseArr = [];
    let memberId = String(this.state.collectionNew[0].memberid);
    for (let link of this.state.collectionNew) {
      let tmp: Array<EventList> = null;
      let selectedVal: any = null;
      tmp = this.state.edetailsIds;
      let scount = link.Count * 10;
      let filteredItem = tmp.filter((i) => i.Id === link.eventid);
      let seventid = String(link.eventid);
      let smemberid = String(link.memberid);
      let sdoe = link.DateOfEvent;
      let stype = link.type;
      let spoints = link.Count * 10;
      let oMember = this.state.membersInfo.filter(x => x.Id.toString() == smemberid)[0];
      let sMemberName = oMember.FirstName + ' ' + oMember.LastName;
      let seventName = this.state.edetailsIds.filter(x => x.Id.toString() == seventid)[0].Title;

      if (filteredItem.length != 0) {
        scount = link.Count * filteredItem[0].Ecount;
      }
      if (true) {
        const listDefinition: any = {
          Title: stype,
          EventId: seventid,
          MemberId: smemberid,
          DateofEvent: sdoe,
          Count: scount,
          MemberName: sMemberName,
          EventName: seventName
        };

        //create promise array
        promiseArr.push(commonServiceManager.createListItem("Event Track Details", listDefinition));
        this.setState((prevState) => ({
          collectionNew: prevState.collectionNew.filter(
            (d) => d.type === "xxx"
          ),
        }));
      }
    }
    //Wait for all promise statements to execute and return true
    Promise.all(promiseArr).then(
      res => {
        this.getListData(memberId);
        this.props.callBack();
        this.setState({ isShow: false });
      }
    ).catch(err => {
      alert(
        "Response status " +
        err.status +
        " - " +
        err.statusText
      );
    });
    this.setState((prevState) => ({
      collectionNew: prevState.collectionNew.filter((d) => d.eventid != 99191),
    }));
    this.setState({ cb: true });
  }

  private getTrackDetailsData(memberid: any, eventid: any): boolean {
    let flag = false;
    this.props.context.spHttpClient
      .get(

        "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('Event Track Details')/Items",
        SPHttpClient.configurations.v1
      )
      .then(async (response: SPHttpClientResponse) => {
        if (response.status === 200) {
          await response.json().then((responseJSON: any) => {
            let i = 0;
            if (responseJSON.value != undefined) {
              while (i < responseJSON.value.length) {
                if (responseJSON.value[i].MemberId == memberid) {
                  if (responseJSON.value[i].EventId == eventid) return flag;
                }
                i++;
              }
            }
          });
        }
      });
    return flag;
  }

  private async getListData(memberId: any): Promise<any> {
    this.setState({ collection: [] });
    const response = await this.props.context.spHttpClient.get(
      "/" +
      this.state.inclusionpath +
      "/" +
      this.state.sitename +
      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000&$orderby=DateofEvent desc",
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      await response.json().then((responseJSON: any) => {
        let i = 0;
        while (i < responseJSON.value.length) {
          if (responseJSON.value[i].MemberId == memberId) {
            if (responseJSON.value[i].MemberId == memberId)
              this.setState((prevState) => ({
                collection: prevState.collection.filter(
                  (d) => d.memberid == memberId
                ),
              }));
            let c = {
              id: i,
              type: responseJSON.value[i].Title,
              eventid: responseJSON.value[i].EventId,
              memberid: memberId,
              Count: responseJSON.value[i].Count,
              DateOfEvent: responseJSON.value[i].DateofEvent,
              MemberName: responseJSON.value[i].MemberName,
              EventName: responseJSON.value[i].eventName

            };
            const newBag = this.state.collection.concat(c);
            this.setState({
              collection: newBag,
              eventid: 0,
            });
          }
          i++;
        }
      });
    }
  }

  public renderFormateDate(collection: any) {
    const formateDateCollection = collection.map((item: any) => {
      return {
        ...item,
        DateOfEvent: moment(item.DateOfEvent).format("MMMM Do, YYYY"),
      };
    });
    return formateDateCollection;
  }

  //Get Member ID of the current user and the Event Track details from Member List 
  public getMemberId(): number {
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
                response.json().then((datada) => {
                  let memberDataIds = datada.value.find(
                    (d: { Title: string }) =>
                      d.Title.toLowerCase() === datauser.Email.toLowerCase()
                  ).ID;
                  this.setState({ newMemberId: memberDataIds });
                  this.setState({ collection: [], membersInfo: datada.value });
                  this.props.context.spHttpClient
                    .get(

                      "/" +
                      this.state.inclusionpath +
                      "/" +
                      this.state.sitename +
                      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000&$orderby=DateofEvent desc",
                      SPHttpClient.configurations.v1
                    )
                    .then((response1: SPHttpClientResponse) => {
                      response1.json().then((responseJSON: any) => {
                        let i = 0;
                        let memberid = localStorage["memberid"];
                        if (
                          memberid === null ||
                          memberid === "undefined" ||
                          memberid === "undefine"
                        ) {
                          memberid = memberDataIds;
                        }
                        while (i < responseJSON.value.length) {
                          if (responseJSON.value[i].MemberId == memberid) {
                            if (responseJSON.value[i].MemberId == memberid)
                              this.setState((prevState) => ({
                                collection: prevState.collection.filter(
                                  (d) => d.memberid == memberid
                                ),
                              }));
                            let c = {
                              id: i,
                              type: responseJSON.value[i].Title,
                              eventid: responseJSON.value[i].EventId,
                              memberid: memberid,
                              Count: responseJSON.value[i].Count,
                              DateOfEvent: responseJSON.value[i].DateofEvent,
                              MemberName: responseJSON.value[i].MemberName,
                              EventName: responseJSON.value[i].eventName

                            };
                            const newBag = this.state.collection.concat(c);
                            this.setState({
                              collection: newBag,
                              eventid: 0,
                            });
                          }
                          i++;
                        }
                      });
                    });
                  return memberDataIds;
                });
              });
          }
        });
      });
    return 0;
  }

  public async componentDidMount() {
    setTimeout(() => {
      let memid: number = 0;
      memid = this.getMemberId();
      this.setState({ loading: false });
    }, 3000);
  }

  public handleSelect = (evt: any) => {
    let ca: string = evt.target.outerText;
    let tmp: Array<EventList> = null;
    let selectedVal: any = null;
    tmp = this.state.edetailsIds;
    let filteredItem = tmp.filter((i) => i.Title.trim() === ca.trim());
    if (filteredItem.length != 0) {
      this.setState({
        selectedkey: 1,
        type: filteredItem[0].Title,
        eventid: filteredItem[0].Id,
        memberid: localStorage["memberid"],
      });
    } else {
      this.setState({
        selectedkey: 0,
        type: ca,
        eventid: 0,
        memberid: localStorage["memberid"],
      });
    }
  }

  private setPoints(e: any): void {
    if (!e.target.value || (e.target.value.length <= 1 && parseInt(e.target.value) <= 5)) {
      this.setState({ points: e.target.value });
    } else {
      this.setState({ points: this.state.points });
    }
  }

  public render() {
    const onRenderCaretDown = (): JSX.Element => {
      return <span></span>;
    };
    return (
      <form>
        <div className="Championview d-flex ">
          {this.state.isShow && <div className="loader"></div>}
          {!this.state.isShow && this.props.showSidebar && (
            <Sidebar
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              becomec={false}
              onClickCancel={() => this.props.onClickCancel()}
              callBack={this.createorupdateItem}
            />
          )}
          <div className="main">
            {this.props.showSidebar && <div className="cv">Championview</div>}
            {!this.state.isShow && (
              <Accordion>
                <Card className="eventsCards">
                  <Accordion.Toggle
                    as={Card.Header}
                    eventKey="0"
                    className="cursor cvw"
                  >
                    {LocaleStrings.ViewDashBoardLabel}
                  </Accordion.Toggle>
                  <Accordion.Collapse eventKey="0">
                    <Card.Body className="cb">
                      <div
                        className="ag-theme-alpine"
                      >
                        <DataGrid
                          rows={this.renderFormateDate(
                            _.orderBy(this.state.collection, ["Id"], ["asc"])
                          )}
                          columns={columns}
                          pageSize={10}
                          loading={this.state.loading}
                        />
                      </div>
                    </Card.Body>
                  </Accordion.Collapse>
                </Card>
                <Card className="eventsCards">
                  <Accordion.Toggle
                    as={Card.Header}
                    eventKey="1"
                    className="cursor"
                  >
                    {LocaleStrings.RecordEventLabel}
                  </Accordion.Toggle>
                  <Accordion.Collapse eventKey="1">
                    <Card.Body className="cb">
                      <div className="form-fields">
                        <div className="form-data">
                          <div className="form-group row">
                            <DatePicker
                              label={LocaleStrings.MonthAndDateLabel}
                              className={`${cx("col-md-4", "col-12", "date")} cv-date-control cv-margin-auto cv-padding-right`}
                              firstDayOfWeek={firstDayOfWeek}
                              strings={DayPickerStrings}
                              showWeekNumbers={true}
                              firstWeekOfYear={1}
                              showMonthPickerAsOverlay={true}
                              placeholder="Select a date..."
                              ariaLabel="Select a date"
                              onSelectDate={this.onChange}
                              value={this.state.DateOfEvent}
                              calloutProps={{ className: "cvDatePickerCallout" }}
                              calendarProps={{ className: "calendarProps", strings: null }}
                              styles={{ callout: { selectors: { '& .ms-DatePicker-day--outfocus': { color: "#757575" } } } }}
                            />
                            <div
                              className={`${cx("col-md-5")} cv-margin-auto cv-padding-right`}>
                              <Dropdown
                                label={LocaleStrings.EventTypeGridLabel}
                                placeholder={LocaleStrings.EventTypeGridLabelPlaceHolder}
                                onChange={(evt) => this.handleSelect(evt)}
                                options={this.options()}
                                onRenderCaretDown={onRenderCaretDown}
                                styles={{ title: { color: "#757575" } }}
                                calloutProps={{ className: "cvEventTypeDropdown" }}
                              />
                            </div>
                            <div className={`${cx("col-md-2")} cv-margin-auto cv-padding-right`}>
                              <TextField
                                label={LocaleStrings.CountLabel}
                                value={this.state.points.toString()}
                                onChange={this.setPoints}
                                id="inputPoints"
                                type="number"
                                min="1"
                                max="5"
                              />
                            </div>
                            <div className={`${cx("col-md-1")} cv-margin-top-auto`}>
                              <Icon iconName="CircleAdditionSolid" className="AddEventIcon"
                                onClick={(_e) =>
                                  this.addDevice(
                                    {
                                      id: this.state.eventUniqueID + 1,
                                      type: this.state.type,
                                      eventid: this.state.eventid,
                                      memberid: this.state.memberid,
                                      Count: this.state.points,
                                      DateOfEvent: this.state.DateOfEvent,
                                      MemberName: "test",
                                      EventName: "evtest",
                                    },
                                    "false"
                                  )
                                } />
                            </div>
                          </div>
                          <div>
                            {this.state.showValidationError &&
                              <span className="errorMessage">
                                {this.state.validationError}
                              </span>
                            }
                          </div>
                          <div className="row">
                            <div>
                              {this.state.collectionNew.map((item) => (
                                <Row className="mt-5 row-margin record-events-grid" key={item.id}>
                                  <Col className="tick" sm={1} xs={1}>
                                    <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className="tickImage" />
                                  </Col>
                                  <Col sm={4} xs={4}>{item.DateOfEvent.toDateString()}</Col>
                                  <Col sm={4} xs={4}>{item.type}</Col>
                                  <Col sm={2} xs={2}>{item.Count}</Col>
                                  <Col sm={1} xs={1}>
                                    <div className="deleteEvent">
                                      <Icon iconName="Delete"
                                        onClick={() => {
                                          this.removeDevice(item.Count, item.id);
                                        }}
                                      />
                                    </div>
                                  </Col>
                                </Row>
                              ))}
                              <br />
                              {this.state.collectionNew !== null &&
                                this.state.collectionNew.length !== 0 && (
                                  <div className="mb-3 helpText">
                                    {LocaleStrings.EventsSubmitMessage}
                                  </div>
                                )}
                              {this.state.collectionNew !== null &&
                                this.state.collectionNew.length !== 0 && (
                                  <DefaultButton
                                    text={LocaleStrings.SubmitButton}
                                    className="mt-4 float-end btnSubmit"
                                    onClick={this.createorupdateItem}
                                  />
                                )}
                            </div>
                          </div>
                        </div>
                      </div>
                    </Card.Body>
                  </Accordion.Collapse>
                </Card>
              </Accordion>
            )}
          </div>
        </div>
      </form>
    );
  }
}
