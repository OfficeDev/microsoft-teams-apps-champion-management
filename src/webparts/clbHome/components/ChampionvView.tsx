import { IDropdownOption } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Label } from '@fluentui/react/lib/Label';
import {
  DatePicker,
  DayOfWeek, IDatePickerStrings
} from "office-ui-fabric-react/lib/DatePicker";
import BootstrapTable from 'react-bootstrap-table-next';
import paginationFactory from 'react-bootstrap-table2-paginator';
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
import { default as commonServices, default as CommonServices } from '../Common/CommonServices';
import siteconfig from "../config/siteconfig.json";
import * as stringConstants from "../constants/strings";
import "../scss/Championview.scss";

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
  callBack?: Function;
  siteUrl: string;
}
export interface ChampionViewState {
  siteUrl: string;
  type: string;
  validationError: string;
  eventid: number;
  memberid: number;
  points: number;
  DateOfEvent: Date;
  collection: Array<ChampList>;
  orderedCollection: Array<ChampList>;
  collectionNew: Array<ChampList>;
  edetails: Array<string>;
  edetailsIds: Array<EventList>;
  selectedkey: string | number;
  isShow: boolean;
  sitename: string;
  inclusionpath: string;
  membersInfo: Array<any>;
  showValidationError: boolean;
  eventUniqueID: number;
  isMultilineNotes: boolean;
  notes: string;
  eventApprovalConfig: string;
  EventsSubmissionMessage: string;
}
export interface ChampList {
  id: number;
  type: string;
  eventid: number;
  memberid: number;
  Count: number;
  DateOfEvent: Date;
  Notes: string;
  MemberName: string;
  EventName: string;
}
export interface EventList {
  Title: string;
  Id: number;
  Ecount: number;
}
export default class ChampionvView extends Component<ChampionViewProps, ChampionViewState>
{
  constructor(props: ChampionViewProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any,
    });

    //Bind Methods
    this.getEventApprovalSetting = this.getEventApprovalSetting.bind(this);
    this.onDateChange = this.onDateChange.bind(this);
    this.getListData = this.getListData.bind(this);
    this.setPoints = this.setPoints.bind(this);
    this.createorupdateItem = this.createorupdateItem.bind(this);
    this.options = this.options.bind(this);
    this.removeDevice = this.removeDevice.bind(this);
    this.getMemberId = this.getMemberId.bind(this);
    this.onNotesChange = this.onNotesChange.bind(this);

    //Class State object
    this.state = {
      siteUrl: this.props.siteUrl,
      type: "",
      validationError: "",
      eventid: 0,
      memberid: 0,
      points: 1,
      DateOfEvent: new Date(),
      collection: [],
      orderedCollection: [],
      collectionNew: [],
      edetails: [],
      edetailsIds: [],
      selectedkey: "",
      isShow: false,
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      membersInfo: [],
      showValidationError: false,
      eventUniqueID: 0,
      isMultilineNotes: false,
      notes: "",
      eventApprovalConfig: "",
      EventsSubmissionMessage: ""
    };
  }

  //Get Member ID of the current user
  public async componentDidMount() {
    await this.getMemberId();
  }

  //Sort the collection by Id column whenever the component is updated
  public componentDidUpdate(_prevProps: Readonly<ChampionViewProps>, prevState: Readonly<ChampionViewState>) {
    if (prevState.collection !== this.state.collection) {
      const orderedCollection = _.orderBy(this.state.collection, ["Id"], ["asc"]);
      this.setState({ orderedCollection: orderedCollection });
    }
  }

  //Get approval status from config List
  public async getEventApprovalSetting() {
    let commonServiceManager = new CommonServices(
      this.props.context,
      this.props.siteUrl
    );
    let filterQuery = "Title eq '" + stringConstants.ChampionEventApprovals + "'";
    const configList: any[] = await commonServiceManager.getItemsWithOnlyFilter(stringConstants.ConfigList, filterQuery);
    this.setState({ eventApprovalConfig: configList[0].Value });
  }

  //Method to be called when date is changed from date picker
  public onDateChange(d: any) {
    this.setState({ DateOfEvent: d });
  }

  //When a new event is added modify the collection to show in the grid
  public addDevice(data: ChampList) {
    if ((data.type == "" || data.type == LocaleStrings.EventTypeGridLabelPlaceHolder)) {
      this.setState({ showValidationError: true, validationError: LocaleStrings.EventTypeValidationMessage, EventsSubmissionMessage: "" });
    }
    else if ((data.Count > 5 || data.Count < 1)) {
      this.setState({ showValidationError: true, validationError: LocaleStrings.CountValidationMessage, EventsSubmissionMessage: "" });
    }
    else {
      this.setState({ collectionNew: [], showValidationError: false, validationError: "", eventUniqueID: data.id });
      const memberEventsData = this.state.collectionNew.concat(data);

      this.setState({
        EventsSubmissionMessage: "",
        collectionNew: memberEventsData,
        points: 1,
        notes: "",
        DateOfEvent: new Date(),
        selectedkey: stringConstants.SelectEventType,
        type: ""
      });
    }
  }

  //Method for dropdown options
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
    myOptions.push({ key: stringConstants.SelectEventType, text: LocaleStrings.EventTypeGridLabelPlaceHolder });
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

  //Method to be called when events are submitted
  public async createorupdateItem() {
    await this.getEventApprovalSetting();
    this.setState({
      isShow: true,
      isMultilineNotes: false,
      eventUniqueID: 0
    });
    let commonServiceManager: commonServices = new commonServices(this.props.context, this.props.siteUrl);
    let promiseArr = [];
    let memberId = String(this.state.collectionNew[0].memberid);
    for (let link of this.state.collectionNew) {
      let tmp: Array<EventList> = null;
      tmp = this.state.edetailsIds;
      let scount = link.Count * 10;
      let filteredItem = tmp.filter((i) => i.Id === link.eventid);
      let seventid = String(link.eventid);
      let smemberid = String(link.memberid);
      let sdoe = link.DateOfEvent;
      let stype = link.type;
      let oMember = this.state.membersInfo.filter(x => x.Id.toString() == smemberid)[0];
      let sMemberName = oMember.FirstName + ' ' + oMember.LastName;
      let seventName = this.state.edetailsIds.filter(x => x.Id.toString() == seventid)[0].Title;
      let sNotes = link.Notes;

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
          EventName: seventName,
          Notes: sNotes,
          Status: this.state.eventApprovalConfig === stringConstants.EnabledStatus ? stringConstants.pendingStatus : stringConstants.approvedStatus
        };

        //create promise array
        promiseArr.push(commonServiceManager.createListItem(stringConstants.EventTrackDetailsList, listDefinition));
        this.setState({
          collectionNew: [],
          showValidationError: false,
          validationError: "",
          EventsSubmissionMessage: (this.state.eventApprovalConfig === stringConstants.EnabledStatus ?
            LocaleStrings.EventSubmissionPendingMessage :
            LocaleStrings.EventsSubmissionSuccessMessage
          )
        });
      }
    }
    //Wait for all promise statements to execute and return true
    Promise.all(promiseArr).then(
      async () => {
        await this.getListData(memberId);
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
    this.setState({ collectionNew: [] });
  }

  //Method for getting events from the sharepoint list
  private async getListData(memberId: any): Promise<any> {
    this.setState({ collection: [] });
    const response = await this.props.context.spHttpClient.get(
      "/" +
      this.state.inclusionpath +
      "/" +
      this.state.sitename +
      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000&$orderby=DateofEvent desc&$filter= Status eq 'Approved' or Status eq null or Status eq ''",
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      await response.json().then((responseJSON: any) => {
        let i = 0;
        const memberEventsData = [];
        while (i < responseJSON.value.length) {
          if (responseJSON.value[i].MemberId == memberId) {
            let eventsData = {
              id: i,
              type: responseJSON.value[i].Title,
              eventid: responseJSON.value[i].EventId,
              memberid: memberId,
              Count: responseJSON.value[i].Count,
              DateOfEvent: new Date(responseJSON.value[i].DateofEvent),
              MemberName: responseJSON.value[i].MemberName,
              EventName: responseJSON.value[i].eventName,
              Notes: responseJSON.value[i].Notes
            };
            memberEventsData.push(eventsData);
          }
          i++;
        }
        this.setState({
          collection: memberEventsData,
          eventid: 0,
        });
      });
    }
  }

  //Get Member ID of the current user and the Event Track details from Member List 
  public async getMemberId() {
    await this.props.context.spHttpClient
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

                  this.setState({ membersInfo: datada.value });
                  this.props.context.spHttpClient
                    .get(

                      "/" +
                      this.state.inclusionpath +
                      "/" +
                      this.state.sitename +
                      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000&$orderby=DateofEvent desc&$filter= Status eq 'Approved' or Status eq null or Status eq ''",
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
                        let memberEventsData: Array<ChampList> = [];
                        while (i < responseJSON.value.length) {
                          if (responseJSON.value[i].MemberId == memberid) {
                            let eventsData: ChampList = {
                              id: i,
                              type: responseJSON.value[i].Title,
                              eventid: responseJSON.value[i].EventId,
                              memberid: memberid,
                              Count: responseJSON.value[i].Count,
                              DateOfEvent: new Date(responseJSON.value[i].DateofEvent),
                              MemberName: responseJSON.value[i].MemberName,
                              EventName: responseJSON.value[i].eventName,
                              Notes: responseJSON.value[i].Notes
                            };
                            memberEventsData.push(eventsData);
                          }
                          i++;
                        }
                        this.setState({
                          collection: memberEventsData,
                          eventid: 0,
                        });
                      });
                    });
                });
              });
          }
        });
      });
  }

  //Method for event type selection from dropdown
  public handleSelect = (_evt: any, item: IDropdownOption) => {
    let ca: string = item.text;
    let tmp: Array<EventList> = null;
    tmp = this.state.edetailsIds;
    let filteredItem = tmp.filter((i) => i.Title.trim() === ca.trim());
    if (filteredItem.length != 0) {
      this.setState({
        selectedkey: item.key,
        type: filteredItem[0].Title,
        eventid: filteredItem[0].Id,
        memberid: localStorage["memberid"],
      });
    } else {
      this.setState({
        selectedkey: item.key,
        type: ca,
        eventid: 0,
        memberid: localStorage["memberid"],
      });
    }
  }

  //Method to set event points
  private setPoints(e: any): void {
    if (!e.target.value || (e.target.value.length <= 1 && parseInt(e.target.value) <= 5)) {
      this.setState({ points: e.target.value });
    } else {
      this.setState({ points: this.state.points });
    }
  }

  //Method to be called when notes text is changed in text field control
  private onNotesChange(_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void {
    const newMultiline = newText.length > stringConstants.NotesMinCharacterLimit;
    if (newMultiline) this.setState({ isMultilineNotes: newMultiline });
    else this.setState({ isMultilineNotes: newMultiline });
    this.setState({ notes: newText });
  }

  //Set pagination properties
  private pagination = paginationFactory({
    page: 1,
    sizePerPage: 10,
    nextPageText: '>',
    prePageText: '<',
    showTotal: true,
    alwaysShowAllBtns: true,
    withFirstAndLast: false,
    hideSizePerPage: true,
    paginationSize: 1
  });


  //Component Render Method
  public render() {

    //Method to be called for not displaying chevron down icon in events type dropdown
    const onRenderCaretDown = (): JSX.Element => {
      return <span></span>;
    };
    const eventsColumns = [
      {
        dataField: "DateOfEvent",
        text: LocaleStrings.DateofEventGridLabel,
        sort: true,
        formatter: (_cell: any, gridRow: any) => <>{gridRow.DateOfEvent.toDateString().slice(4)}</>
      },
      {
        dataField: "type",
        text: LocaleStrings.EventTypeGridLabel,
        sort: true
      },
      {
        dataField: "Count",
        text: LocaleStrings.CMPSideBarPointsLabel,
        sort: true
      },

    ];

    return (
      <form>
        <div className="Championview d-flex ">
          {this.state.isShow && <div className="loader"></div>}
          <div className="main">
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
                    <div className="events-table">
                      <BootstrapTable
                        bootstrap4
                        keyField="id"
                        data={this.state.orderedCollection}
                        columns={eventsColumns}
                        striped
                        headerWrapperClasses="event-table-header"
                        hover
                        table-responsive={true}
                        pagination={this.pagination}
                        hidePageListOnlyOnePage={true}
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
                            className={`${cx("col-md-3", "col-12", "date")} cv-date-control cv-margin-auto cv-padding-right`}
                            firstDayOfWeek={firstDayOfWeek}
                            strings={DayPickerStrings}
                            showWeekNumbers={true}
                            firstWeekOfYear={1}
                            showMonthPickerAsOverlay={true}
                            ariaLabel={LocaleStrings.SelectDate}
                            onSelectDate={this.onDateChange}
                            value={this.state.DateOfEvent}
                            calloutProps={{ className: "cvDatePickerCallout" }}
                            calendarProps={{ className: "calendarProps", strings: null }}
                            styles={{ callout: { selectors: { '& .ms-DatePicker-day--outfocus': { color: "#757575" } } } }}
                          />
                          <div
                            className={`${cx("col-md-3")} cv-margin-auto cv-padding-right`}>
                            <Dropdown
                              label={LocaleStrings.EventTypeGridLabel}
                              placeholder={LocaleStrings.EventTypeGridLabelPlaceHolder}
                              onChange={(evt, item) => this.handleSelect(evt, item)}
                              options={this.options()}
                              onRenderCaretDown={onRenderCaretDown}
                              styles={{ title: { color: "#757575" } }}
                              calloutProps={{ className: "cvEventTypeDropdown" }}
                              selectedKey={this.state.selectedkey}
                            />
                          </div>
                          <div className={`${cx("col-md-2")} cv-margin-auto cv-padding-right`}>
                            <TextField
                              label={LocaleStrings.CountLabel}
                              value={this.state.points.toString()}
                              onChange={this.setPoints}
                              type="number"
                              min="1"
                              max="5"
                            />
                          </div>
                          <div className={`${cx("col-md-3")} ${this.state.isMultilineNotes ? "cv-notes-margin-top" : "cv-margin-auto"} cv-padding-right`}>
                            <TextField
                              label={LocaleStrings.NotesLabel}
                              placeholder={LocaleStrings.NotesPlaceholder}
                              multiline={this.state.isMultilineNotes}
                              onChange={this.onNotesChange}
                              maxLength={200}
                              value={this.state.notes}
                            />
                          </div>
                          <div className={`${cx("col-md-1")} ${this.state.isMultilineNotes ? "cv-add-event-margin-top" : "cv-margin-top-auto"}`}>
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
                                    Notes: this.state.notes,
                                    MemberName: "test",
                                    EventName: "evtest",
                                  }
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
                        {this.state.EventsSubmissionMessage !== "" ?
                          <Label className="pendingEventsMessage">
                            <img src={require('../assets/TOTImages/tickIcon.png')} alt={LocaleStrings.SuccessIcon} className="tickImage" />
                            {this.state.EventsSubmissionMessage}
                          </Label>
                          : null
                        }
                        <div className="row">
                          <div>
                            {this.state.collectionNew.map((item) => (
                              <Row className="mt-5 row-margin record-events-grid" key={item.id}>
                                <Col className="tick" sm={1} xs={1}>
                                  <img src={require('../assets/TOTImages/tickIcon.png')} alt={LocaleStrings.SuccessIcon} className="tickImage" />
                                </Col>
                                <Col sm={3} xs={3} className="cv-events-text-align">{item.DateOfEvent.toDateString()}</Col>
                                <Col sm={3} xs={3} className="cv-events-text-align">{item.type}</Col>
                                <Col sm={1} xs={1} className="cv-events-text-align">{item.Count}</Col>
                                <Col sm={3} xs={3} className="cv-events-text-align">{item.Notes}</Col>
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
          </div>
        </div>
      </form>
    );
  }
}
