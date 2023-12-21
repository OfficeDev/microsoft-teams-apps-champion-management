import { Dialog, DialogType, IDropdownOption, PrimaryButton, Spinner } from '@fluentui/react';
import { Icon } from '@fluentui/react/lib/Icon';
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import _ from "lodash";
import {
  DatePicker,
  DayOfWeek, IDatePickerStrings
} from "office-ui-fabric-react/lib/DatePicker";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import React, { Component } from "react";
import BootstrapTable from 'react-bootstrap-table-next';
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import { default as commonServices, default as CommonServices } from '../Common/CommonServices';
import siteconfig from "../config/siteconfig.json";
import * as stringConstants from "../constants/strings";
import "../scss/RecordEvents.scss";

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
export interface RecordEventsProps {
  context: WebPartContext;
  callBack?: Function;
  siteUrl: string;
  showRecordEventPopup: boolean;
  updateRecordEventsPopupState: Function;
  setEventsSubmissionMessage: Function;
  currentThemeName?: string;
}
export interface RecordEventsState {
  siteUrl: string;
  type: string;
  validationError: string;
  eventid: number;
  memberid: number;
  points: number;
  DateOfEvent: Date;
  collectionNew: Array<ChampList>;
  edetails: Array<string>;
  edetailsIds: Array<EventList>;
  selectedkey: string | number;
  showLoader: boolean;
  sitename: string;
  inclusionpath: string;
  membersInfo: Array<any>;
  showValidationError: boolean;
  eventUniqueID: number;
  isMultilineNotes: boolean;
  notes: string;
  eventApprovalConfig: string;
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
export default class RecordEvents extends Component<RecordEventsProps, RecordEventsState>
{
  constructor(props: RecordEventsProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context as any,
    });

    //Bind Methods
    this.getEventApprovalSetting = this.getEventApprovalSetting.bind(this);
    this.onDateChange = this.onDateChange.bind(this);
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
      collectionNew: [],
      edetails: [],
      edetailsIds: [],
      selectedkey: "",
      showLoader: false,
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      membersInfo: [],
      showValidationError: false,
      eventUniqueID: 0,
      isMultilineNotes: false,
      notes: "",
      eventApprovalConfig: ""
    };
  }

  //Get Member ID of the current user
  public async componentDidMount() {
    await this.getMemberId();
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
      this.setState({ showValidationError: true, validationError: LocaleStrings.EventTypeValidationMessage });
      this.props.setEventsSubmissionMessage("");
    }
    else if ((data.Count > 5 || data.Count < 1)) {
      this.setState({ showValidationError: true, validationError: LocaleStrings.CountValidationMessage });
      this.props.setEventsSubmissionMessage("");
    }
    else {
      this.setState({ collectionNew: [], showValidationError: false, validationError: "", eventUniqueID: data.id });
      const memberEventsData = this.state.collectionNew.concat(data);
      this.props.setEventsSubmissionMessage("");
      this.setState({
        collectionNew: memberEventsData,
        points: 1,
        notes: "",
        DateOfEvent: new Date(),
        selectedkey: stringConstants.SelectEventType,
        type: "",
        isMultilineNotes: false
      });
    }
  }

  //Method for dropdown options
  public options() {
    let optionArray: any = [];
    let optionArrayIds: any = [];
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
      showLoader: true,
      eventUniqueID: 0
    });
    let commonServiceManager: commonServices = new commonServices(this.props.context, this.props.siteUrl);
    let promiseArr = [];
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

      if (filteredItem.length !== 0) {
        scount = link.Count * filteredItem[0].Ecount;
      }
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
    }
    this.setState({
      collectionNew: [],
      showValidationError: false,
      validationError: ""
    });
    //Wait for all promise statements to execute and return true
    Promise.all(promiseArr).then(
      async () => {
        this.props.callBack();
        this.setState({ showLoader: false });
        this.props.setEventsSubmissionMessage(
          this.state.eventApprovalConfig === stringConstants.EnabledStatus ?
            LocaleStrings.EventSubmissionPendingMessage :
            LocaleStrings.EventsSubmissionSuccessMessage
        );
        this.props.updateRecordEventsPopupState(false);
      }
    ).catch(err => {
      alert(
        "Response status " +
        err.status +
        " - " +
        err.statusText
      );
    });
    this.setState({ collectionNew: [], eventid: 0 });
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
                response.json().then((datada: any) => {
                  let memberDataIds = datada.value.find(
                    (d: { Title: string }) =>
                      d.Title.toLowerCase() === datauser.Email.toLowerCase()
                  ).ID;
                  this.setState({ membersInfo: datada.value, eventid: 0 });
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
    if (filteredItem.length !== 0) {
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

  //Component Render Method
  public render() {

    //Method to be called for not displaying chevron down icon in events type dropdown
    const onRenderCaretDown = (): JSX.Element => {
      return <span></span>;
    };

    const eventsTableHeader: any = [
      {
        dataField: "id",
        headerAttrs: { hidden: true },
        title: true,
        formatter: () => <img src={require('../assets/TOTImages/tickIcon.png')}
          alt={LocaleStrings.SuccessIcon} className="tickImage" />
      },
      {
        dataField: "DateOfEvent",
        headerAttrs: { hidden: true },
        title: (_cell: any, gridRow: any) => gridRow.DateOfEvent.toDateString(),
        formatter: (_: any, gridRow: any) => <>{gridRow.DateOfEvent.toDateString()}</>
      },
      {
        dataField: "type",
        headerAttrs: { hidden: true },
        title: true,
      },
      {
        dataField: "Count",
        headerAttrs: { hidden: true },
        title: true,
      },

      {
        dataField: "Notes",
        headerAttrs: { hidden: true },
        title: true,
      },
      {
        dataField: "Remove",
        headerAttrs: { hidden: true },
        title: () => LocaleStrings.RemoveEventLabel,
        formatter: (_: any, gridRow: any) =>
          <Icon iconName="Delete"
            onClick={() => {
              this.removeDevice(gridRow.Count, gridRow.id);
            }}
            className="event-delete-icon"
          />
      }
    ];
    const isDarkOrContrastTheme = this.props.currentThemeName === stringConstants.themeDarkMode || this.props.currentThemeName === stringConstants.themeContrastMode;
    return (
      <Dialog
        hidden={!this.props.showRecordEventPopup}
        onDismiss={() => !this.state.showLoader && this.props.updateRecordEventsPopupState(false)}
        modalProps={{
          isBlocking: true,
          className: `record-events-popup${isDarkOrContrastTheme ? " record-events-popup-" + this.props.currentThemeName : ""}`
        }}
        dialogContentProps={{ type: DialogType.normal, title: LocaleStrings.RecordEventLabel, className: "dialog-content" }}
      >
        <>
          <Row>
            <Col xl={3} lg={3} md={12} sm={12}>
              <DatePicker
                label={LocaleStrings.MonthAndDateLabel}
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
                className="record-events-date-picker"
              />
            </Col>
            <Col xl={3} lg={3} md={12} sm={12}>
              <Dropdown
                label={LocaleStrings.EventTypeGridLabel}
                placeholder={LocaleStrings.EventTypeGridLabelPlaceHolder}
                onChange={(evt, item) => this.handleSelect(evt, item)}
                options={this.options()}
                onRenderCaretDown={onRenderCaretDown}
                calloutProps={{ className: "cvEventTypeDropdown" }}
                selectedKey={this.state.selectedkey}
                className="record-events-dropdown"
              />
            </Col>
            <Col xl={2} lg={2} md={12} sm={12}>
              <TextField
                label={LocaleStrings.CountLabel}
                value={this.state.points.toString()}
                onChange={this.setPoints}
                type="number"
                min="1"
                max="5"
                className='record-events-count-field'
              />
            </Col>
            <Col xl={3} lg={3} md={12} sm={12}>
              <TextField
                label={LocaleStrings.NotesLabel}
                placeholder={LocaleStrings.NotesPlaceholder}
                multiline={this.state.isMultilineNotes}
                onChange={this.onNotesChange}
                maxLength={200}
                value={this.state.notes}
                className='record-events-notes-field'
              />
            </Col>
            <Col xl={1} lg={1} md={12} sm={12}>
              <Icon iconName="CircleAdditionSolid" className="add-event-icon" title={LocaleStrings.AddEventLabel}
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
            </Col>
          </Row>
          {this.state.showValidationError &&
            <div className="error-message">
              {this.state.validationError}
            </div>
          }
          {this.state.collectionNew.length > 0 &&
            <BootstrapTable
              striped
              bootstrap4
              keyField="id"
              data={this.state.collectionNew}
              columns={eventsTableHeader}
              table-responsive={true}
              wrapperClasses="events-collection-table"
            />
          }
          {this.state.showLoader &&
            <Spinner
              label={LocaleStrings.ProcessingSpinnerLabel}
              className='submission-spinner'
              ariaLive="assertive" labelPosition="left" />
          }
          {this.state.collectionNew !== null &&
            this.state.collectionNew.length !== 0 && (
              <div className="mb-3 help-text">
                {LocaleStrings.EventsSubmitMessage}
              </div>
            )}
          {this.state.collectionNew !== null &&
            this.state.collectionNew.length !== 0 && (
              <PrimaryButton
                text={LocaleStrings.SubmitButton}
                className="mt-3 mb-3 record-events-submit-btn"
                onClick={this.createorupdateItem}
              />
            )}
        </>
      </Dialog>
    );
  }
}
