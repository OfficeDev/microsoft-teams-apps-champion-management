import React, { Component } from "react";
import "../scss/Championview.scss";
import Sidebar from "../components/Sidebar";
import {
  DatePicker,
  DayOfWeek,
  IDatePickerStrings,
} from "office-ui-fabric-react/lib/DatePicker";
import { mergeStyleSets } from "office-ui-fabric-react/lib/Styling";
import cx from "classnames";
import { DefaultButton } from "office-ui-fabric-react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Dropdown, IDropdown } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { sp } from "@pnp/sp";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import Accordion from "react-bootstrap/Accordion";
import Card from "react-bootstrap/Card";
import Alert from "@material-ui/lab/Alert";
import { DataGrid } from "@material-ui/data-grid";
import * as moment from "moment";
import siteconfig from "../config/siteconfig.json";
import _ from "lodash";

const columns = [
  { field: "DateOfEvent", headerName: "Date of Event", width: 250 },
  { field: "type", headerName: "Type", width: 250 },
  { field: "Count", headerName: "Points", width: 250 },
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

const controlClass = mergeStyleSets({
  control: {
    margin: "0 0 15px 0",
    maxWidth: "300px",
  },
});

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
  membersInfo: Array<any>
}
export interface ChampList {
  id: number;
  type: string;
  eventid: number;
  memberid: number;
  Count: number;
  DateOfEvent: Date;
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
      membersInfo: []
    };
  }

  public onChange(d: any) {
    this.setState({ DateOfEvent: d });
  }

  public addDevice(data: ChampList, saved: any) {
    if (saved === "false") {
      this.setState({ collectionNew: [] });
      const newBag = this.state.collectionNew.concat(data);
      this.setState({
        collectionNew: newBag,
        eventid: 0,
        points: data.Count,
      });
      this.setState({ selectedkey: 0 });
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

  public removeDevice(type: string, points: number) {
    this.setState((prevState) => ({
      collectionNew: prevState.collectionNew.filter(
        (data) => data.type !== type,
        points
      ),
    }));
  }

  public createorupdateItem() {
    this.setState({ isShow: true });
    for (let link of this.state.collectionNew) {
      let tmp: Array<EventList> = null;
      let selectedVal: any = null;
      tmp = this.state.edetailsIds;
      let scount = link.Count * 10;
      let item1 = tmp.filter((i) => i.Id === link.eventid);
      let seventid = String(link.eventid);
      let smemberid = String(link.memberid);
      let sdoe = link.DateOfEvent;
      let stype = link.type;
      let spoints = link.Count * 10;
      let oMember = this.state.membersInfo.filter(x => x.Id.toString() == smemberid)[0];
      let sMemberName = oMember.FirstName + ' ' + oMember.LastName;
      let seventName = this.state.edetailsIds.filter(x => x.Id.toString() == seventid)[0].Title;

      if (item1.length != 0) {
        scount = link.Count * item1[0].Ecount;
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

        const spHttpClientOptions: ISPHttpClientOptions = {
          body: JSON.stringify(listDefinition),
        };

        if (true) {
          setTimeout(() => {
            this.setState({ isShow: false });
          }, 2000);

          this.props.callBack();

          const url: string =
            "/" +
            this.state.inclusionpath +
            "/" +
            this.state.sitename +
            "/_api/web/lists/GetByTitle('Event Track Details')/items";
          if (this.props.context)
            this.props.context.spHttpClient
              .post(
                url,
                SPHttpClient.configurations.v1,
                spHttpClientOptions
              )
              .then((responseData: SPHttpClientResponse) => {
                this.addDevice(link, "true");
                if (responseData.status === 201) {
                  this.getListData(smemberid, seventid);
                } else {
                  alert(
                    "Response status " +
                    responseData.status +
                    " - " +
                    responseData.statusText
                  );
                }
              })
              .catch((error) => alert(error.message));
        } else {
        }
        this.setState((prevState) => ({
          collectionNew: prevState.collectionNew.filter(
            (d) => d.type === "xxx"
          ),
        }));
      }
    }
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

  private async getListData(memberid: any, _eventid: any): Promise<any> {
    this.setState({ collection: [] });
    const response = await this.props.context.spHttpClient.get(

      "/" +
      this.state.inclusionpath +
      "/" +
      this.state.sitename +
      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000",
      SPHttpClient.configurations.v1
    );
    if (response.status === 200) {
      await response.json().then((responseJSON: any) => {
        let i = 1;
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
    }
  }

  public renderFormateDate(collection: any) {
    const formateDateCollection = collection.map((item: any) => {
      return {
        ...item,
        DateOfEvent: moment(item.DateOfEvent).format("MMMM Do YYYY, h:mm:ss a"),
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
                      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$top=5000",
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
    let item1 = tmp.filter((i) => i.Title === ca);
    if (item1.length != 0) {
      this.setState({
        selectedkey: 1,
        type: item1[0].Title,
        eventid: item1[0].Id,
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
                <Card>
                  <Accordion.Toggle
                    as={Card.Header}
                    eventKey="0"
                    className="cursor cvw"
                  >
                    View Dashboard
                  </Accordion.Toggle>
                  <Accordion.Collapse eventKey="0">
                    <Card.Body className="cb">
                      <div
                        className="ag-theme-alpine"
                        style={{
                          height: 400,
                          width: "auto",
                          backgroundColor: "rgba(158,187,208,.5)",
                        }}
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
                <Card>
                  <Accordion.Toggle
                    as={Card.Header}
                    eventKey="1"
                    className="cursor"
                  >
                    Record Event
                  </Accordion.Toggle>
                  <Accordion.Collapse eventKey="1">
                    <Card.Body className="cb">
                      <div className="form-fields">
                        <div className="form-data">
                          <div className="form-group row">
                            <label
                              htmlFor="date"
                              className="col-sm-3 col-form-label"
                            >
                              Month and Date
                            </label>
                            <DatePicker
                              className={cx(
                                controlClass.control,
                                "col-sm-9",
                                "date"
                              )}
                              firstDayOfWeek={firstDayOfWeek}
                              strings={DayPickerStrings}
                              showWeekNumbers={true}
                              firstWeekOfYear={1}
                              showMonthPickerAsOverlay={true}
                              placeholder="Select a date..."
                              ariaLabel="Select a date"
                              onSelectDate={this.onChange}
                              value={this.state.DateOfEvent}
                            />
                          </div>
                          <div className="form-group row">
                            <label
                              htmlFor="type"
                              className="col-sm-3 col-form-label"
                            >
                              Type
                            </label>
                            <div className="col-sm-9">
                              <Dropdown
                                placeholder="Select Event Type"
                                onChange={(evt) => this.handleSelect(evt)}
                                id="drp"
                                options={this.options()}
                                onRenderCaretDown={onRenderCaretDown}
                              />
                            </div>
                          </div>
                          {this.state.type &&
                            this.state.type !== "Select Event Type" ? (
                            <div className="form-group row">
                              <label
                                htmlFor="inputPoints"
                                className="col-sm-3 col-form-label"
                              >


                                Count
                              </label>
                              <div className="col-sm-9">
                                <TextField
                                  value={this.state.points.toString()}
                                  onChange={this.setPoints}
                                  id="inputPoints"
                                  type="number"
                                  min="1"
                                  max="5"
                                />
                              </div>
                            </div>
                          ) : (
                            ""
                          )}
                          <div className="row">
                            <div className="col-12">
                              <div className="float-end">
                                <DefaultButton
                                  text="Add"
                                  onClick={(_e) =>
                                    this.addDevice(
                                      {
                                        id: 0,
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
                                  }
                                />
                              </div>
                              <br />
                              <br />
                              {this.state.collectionNew !== null &&
                                this.state.collectionNew.length !== 0 && (
                                  <div className="mb-3">
                                    When you are done adding Events, Please
                                    Click on <b>Submit</b> to save
                                  </div>
                                )}

                              {this.state.collectionNew.map((item) => (
                                <div key={item.eventid} className="m-2 mb-3">
                                  <Alert
                                    onClose={() => {
                                      this.removeDevice(item.type, item.Count);
                                    }}
                                  >
                                    {item.type}
                                    <span className="ml-4">{item.Count}</span>
                                  </Alert>
                                </div>
                              ))}

                              {this.state.collectionNew !== null &&
                                this.state.collectionNew.length !== 0 && (
                                  <DefaultButton
                                    text="Submit"
                                    className="mt-4 float-end"
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
