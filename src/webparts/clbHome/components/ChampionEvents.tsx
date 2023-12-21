import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { app } from '@microsoft/teams-js-v2';
import * as React from 'react';
import * as stringsConstants from '../constants/strings';
import BootstrapTable from 'react-bootstrap-table-next';
import paginationFactory from 'react-bootstrap-table2-paginator';
import Col from 'react-bootstrap/esm/Col';
import moment from 'moment';
import Row from 'react-bootstrap/Row';
import { Component } from 'react';
import { Icon, initializeIcons } from 'office-ui-fabric-react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '../scss/Champions.scss';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

initializeIcons();

let currentUserName: string;

interface ChampionEventsProps {
    context: WebPartContext;
    callBack?: Function;
    filteredAllEvents: Array<any>;
    selectedMemberDetails?: any;
    parentComponent: string;
    selectedMemberID: string;
    loggedinUserEmail?: string;
}
interface ChampionEventsState {
    selectedUserActivities: Array<any>;
    selectedMemberDetails: Array<any>;
}

export default class ChampionEvents extends Component<ChampionEventsProps, ChampionEventsState> {
    constructor(props: any) {
        super(props);
        this.state = {
            selectedUserActivities: [],
            selectedMemberDetails: []
        };

        currentUserName = this.props.context.pageContext.user.displayName;
        this._renderListAsync();
    }

    //Initializes the teams library and calling the methods to load the initial data  
    public _renderListAsync() {
        app.initialize();
    }

    //method to load the selected member activities
    public componentDidMount() {
        this.getChampionActivities(this.props.selectedMemberID);
    }
    public componentDidUpdate(prevProps: Readonly<ChampionEventsProps>, prevState: Readonly<ChampionEventsState>): void {
        if (prevState.selectedUserActivities !== this.state.selectedUserActivities && this.state.selectedUserActivities.length > 0) {
            const paginationWrapper = document?.getElementsByClassName("react-bootstrap-table-pagination")[0];
            const divElements = paginationWrapper?.getElementsByTagName('div');
            divElements[0]?.setAttribute("class", "col-md-8 col-xs-8 col-sm-8 col-lg-8");
            divElements[1]?.setAttribute("class", "react-bootstrap-table-pagination-list col-md-4 col-xs-4 col-sm-4 col-lg-4");
        }
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

    //get data for the activities table
    private async getChampionActivities(selectedChampion: string) {
        let memberActivitesArray: any = [];
        let memberDetails: any = [];
        this.setState({
            selectedUserActivities: [],
            selectedMemberDetails: []
        });
        if (selectedChampion !== stringsConstants.AllLabel && selectedChampion !== "") {

            //filtering the selected member's data from the array of all records
            let selectedMemberEvents = this.props.filteredAllEvents.filter((item) => item.MemberId === selectedChampion);

            //creating an array to store the required data for Activities table
            selectedMemberEvents.forEach((event) => {
                memberActivitesArray.push({
                    DateofEvent: event[stringsConstants.dateOfEventLabel],
                    Type: event["EventName"] ? event["EventName"] : "",
                    Points: event["Count"] ? event["Count"] : ""
                });
            });

            //creating an array to store the required data to show member details
            if (this.props.parentComponent === stringsConstants.ChampionReportLabel) {
                memberDetails.push({
                    ID: this.props.selectedMemberDetails[0].ID,
                    Title: this.props.selectedMemberDetails[0].Title,
                    FirstName: this.props.selectedMemberDetails[0].FirstName,
                    LastName: this.props.selectedMemberDetails[0].LastName,
                });
            }
            else if (this.props.parentComponent === stringsConstants.ChampionsCardsLabel) {
                memberDetails.push({
                    Points: this.props.selectedMemberDetails.Points,
                    ID: this.props.selectedMemberDetails.ID,
                    Title: this.props.selectedMemberDetails.Title,
                    FirstName: this.props.selectedMemberDetails.FirstName,
                    LastName: this.props.selectedMemberDetails.LastName,
                    Rank: this.props.selectedMemberDetails.Rank
                });
            }

            this.setState({
                selectedUserActivities: memberActivitesArray,
                selectedMemberDetails: memberDetails
            });
        }
    }

    //render the sort caret on the header column for accessbility issues fix
    customSortCaret = (order: any, column: any) => {
        const ariaLabel = navigator.userAgent.match(/iPhone/i) ? "sortable" : "";
        const id = column.dataField;
        if (!order) {
            return (
                <span className="sort-order" id={id} aria-label={ariaLabel}>
                    <span className="dropdown-caret">
                    </span>
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'asc') {
            if (column.dataField === stringsConstants.dateOfEventLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.typeLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.pointsLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', stringsConstants.sortAscAriaSort);
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', stringsConstants.sortAscAriaSort);
            }

            return (
                <span className="sort-order">
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'desc') {
            if (column.dataField === stringsConstants.dateOfEventLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.typeLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', "");
            } else if (column.dataField === stringsConstants.pointsLabel) {
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-sort', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-sort', stringsConstants.sortDescAriaSort);
                document.getElementById(stringsConstants.dateOfEventId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventTypeId).setAttribute('aria-description', "");
                document.getElementById(stringsConstants.eventPointsId).setAttribute('aria-description', stringsConstants.sortDescAriaSort);
            }

            return (
                <span className="sort-order">
                    <span className="dropdown-caret">
                    </span>
                </span>);
        }
        return null;
    }

    //custom header format for sortable column for accessbility
    headerFormatter(column: any, colIndex: any, { sortElement, filterElement }: any) {
        //adding sortable information to aria-label to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            const id = column.dataField;
            return (
                <button tabIndex={-1} aria-describedby={id} aria-label={column.text} className='sort-header'>
                    {column.text}
                    {sortElement}
                </button>
            );
        }
        else {
            return (
                <div aria-hidden="true" title={column.text} className='header-div-wrapper'>
                    <span className='header-span-text'>{column.text}</span>
                    {sortElement}
                </div>
            );
        }
    }

    // format the cell for Event Date column to fix accessibility issues
    eventDateFormatter = (cell: any) => {
        const ariaLabel = `${LocaleStrings.DateofEventGridLabel} ${cell}`
        return (
            <span aria-label={ariaLabel}>
                <span title={moment(new Date(cell)).format("MMMM Do, YYYY")} aria-hidden="true">
                    {moment(new Date(cell)).format("MMMM Do, YYYY")}
                </span>
            </span>
        );
    }

    // format the cell for Event Type column to fix accessibility issues
    eventTypeFormatter = (cell: any) => {
        const ariaLabel = `${LocaleStrings.EventTypeGridLabel} ${cell}`
        return (
            <span aria-label={ariaLabel}>
                <span title={cell} aria-hidden="true">
                    {cell}
                </span>
            </span>
        );
    }

    // format the cell for Event Points column to fix accessibility issues
    eventPointsFormatter = (cell: any) => {
        const ariaLabel = `${LocaleStrings.CMPSideBarPointsLabel} ${cell}`
        return (
            <span aria-label={ariaLabel}>
                <span title={cell} aria-hidden="true">
                    {cell}
                </span>
            </span>
        );
    }

    //Set pagination properties
    private pagination: any = paginationFactory({
        page: 1,
        sizePerPage: 10,
        paginationSize: 1,
        nextPageText: '>',
        prePageText: '<',
        alwaysShowAllBtns: true,
        withFirstAndLast: false,
        hideSizePerPage: true,
        showTotal: true,
        paginationTotalRenderer: (from: any, to: any, size: any) => {
            const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} Results` : ""
            return (
                <span className="react-bootstrap-table-pagination-total" aria-live="polite" role="alert">
                    {resultsFound}
                </span>
            )
        }
    });

    //Main render method
    public render() {
        // To determine whether the component is called from sidebar or not
        const isSidebar = this.props.parentComponent === stringsConstants.SidebarLabel;
        const isChampionReport = this.props.parentComponent === stringsConstants.ChampionReportLabel;

        const activitiesTableHeader: any = [
            {
                dataField: stringsConstants.dateOfEventLabel,
                formatter: this.eventDateFormatter,
                headerFormatter: this.headerFormatter,
                text: LocaleStrings.DateofEventGridLabel,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.dateOfEventId, "role": "columnheader", "scope": "col" },
                attrs: { 'role': 'presentation' },
                sortValue: (cell: any) => new Date(cell)
            },
            {
                dataField: stringsConstants.typeLabel,
                formatter: this.eventTypeFormatter,
                headerFormatter: this.headerFormatter,
                text: LocaleStrings.EventTypeGridLabel,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.eventTypeId, "role": "columnheader", "scope": "col" },
                attrs: { 'role': 'presentation' }
            },
            {
                dataField: stringsConstants.pointsLabel,
                formatter: this.eventPointsFormatter,
                headerFormatter: this.headerFormatter,
                text: LocaleStrings.CMPSideBarPointsLabel,
                sort: true,
                sortCaret: this.customSortCaret,
                headerAttrs: { "id": stringsConstants.eventPointsId, "role": "columnheader", "scope": "col" },
                attrs: { 'role': 'presentation' },
                sortValue: (cell: any) => parseInt(cell)
            }
        ];

        return (
            <div className="gtc-cards">
                <div className="showActivitiesPopupBody">
                    <Row xl={isSidebar ? 1 : 2} lg={isSidebar ? 1 : 2} md={1} sm={1} xs={1} className="report-profile-grid-wrapper">
                        {this.props.parentComponent !== stringsConstants.SidebarLabel &&
                            <Col xl={isChampionReport ? 3 : 4} lg={isChampionReport ? 3 : 4} md={isChampionReport ? 5 : 12} sm={12} xs={12} className={isChampionReport ? "report-profile-wrapper" : ""}>
                                {this.state.selectedMemberDetails.length > 0 &&
                                    <>
                                        {isChampionReport && <div className='events-profile-heading'>{LocaleStrings.ChampionLabel}</div>}
                                        <div className="showActivitiesImage-IconArea">
                                            <Person
                                                personQuery={this.state.selectedMemberDetails[0].Title}
                                                view={3}
                                                personCardInteraction={1}
                                                verticalLayout={true}
                                                avatarSize='large'
                                                className="activities-profile-image"
                                            />
                                            {this.props.parentComponent === stringsConstants.ChampionsCardsLabel &&
                                                <div className="showActivities-rank-points-block">
                                                    <span className="showActivities-rank" title={`Rank ${this.state.selectedMemberDetails[0].Rank}`}>Rank <span className="showActivities-rank-value"># {this.state.selectedMemberDetails[0].Rank}</span></span>
                                                    <span className="showActivities-points" title={`#${this.state.selectedMemberDetails[0].Points}`}>
                                                        {this.state.selectedMemberDetails[0].Points}
                                                        <Icon iconName="FavoriteStarFill" className="showActivities-points-star" />
                                                    </span>
                                                </div>
                                            }
                                            {this.props.loggedinUserEmail !== this.state.selectedMemberDetails[0].Title &&
                                                <div className="showActivities-icon-area">
                                                    <div className="request-to-call-link"
                                                        title={LocaleStrings.RequestToCallLabel}
                                                        onClick={() => this.openTask("https://teams.microsoft.com/l/meeting/new?subject=" +
                                                            currentUserName + " / " + this.state.selectedMemberDetails[0].FirstName + " " + this.state.selectedMemberDetails[0].LastName + " " + LocaleStrings.MeetupSubject +
                                                            "&content=" + LocaleStrings.MeetupBody + "&attendees=" + this.state.selectedMemberDetails[0].Title)}
                                                    > {LocaleStrings.RequestToCallLabel}
                                                    </div>
                                                </div>
                                            }
                                        </div>
                                    </>
                                }
                            </Col>
                        }
                        <Col xl={isSidebar ? 12 : isChampionReport ? 7 : 8} lg={isSidebar ? 12 : isChampionReport ? 7 : 8} md={isChampionReport ? 7 : 12}
                            sm={12} xs={12} className={isChampionReport ? "report-grid-wrapper" : ""}>
                            {isChampionReport && <div className="events-grid-heading">{LocaleStrings.ActivitiesLabel}</div>}
                            <div className="showActivities-grid-area">
                                {!isChampionReport && <div className="activities-grid-heading">{LocaleStrings.ActivitiesLabel}</div>}
                                <BootstrapTable
                                    bootstrap4
                                    keyField={'dateOfEvents'}
                                    data={this.state.selectedUserActivities}
                                    columns={activitiesTableHeader}
                                    pagination={this.pagination}
                                    table-responsive={true}
                                    wrapperClasses="events-table-wrapper-class"
                                    noDataIndication={() => (<div className='activities-noRecordsFound'>{LocaleStrings.NoActivitiesinGridLabel}</div>)}
                                />
                            </div>
                        </Col>
                    </Row>
                </div>
            </div>
        );
    }
}
