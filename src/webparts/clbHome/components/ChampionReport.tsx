import * as constants from '../constants/strings';
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as stringsConstants from '../constants/strings';
import Card from 'react-bootstrap/esm/Card';
import ChampionEvents from './ChampionEvents';
import Col from 'react-bootstrap/esm/Col';
import commonServices from '../Common/CommonServices';
import EventsChart from './EventsChart';
import moment from 'moment';
import React from 'react';
import Row from 'react-bootstrap/esm/Row';
import styles from '../scss/ChampionReport.module.scss';
import {
    Callout,
    ChoiceGroup,
    ComboBox,
    DefaultButton,
    DirectionalHint,
    IChoiceGroupOption,
    IComboBox,
    IComboBoxOption,
    Icon,
    PrimaryButton
} from '@fluentui/react';
import { DatePicker, defaultDatePickerStrings } from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';

//Global Variables
let commonServiceManager: commonServices;

const options: IChoiceGroupOption[] = [
    { key: stringsConstants.AllTime, text: LocaleStrings.AllTime },
    { key: stringsConstants.LastMonth, text: LocaleStrings.LastMonth },
    { key: stringsConstants.Last3Months, text: LocaleStrings.Last3Months },
    { key: stringsConstants.Last6Months, text: LocaleStrings.Last6Months },
    { key: stringsConstants.Last1Year, text: LocaleStrings.Last1Year },
    { key: stringsConstants.SelectaCustomDate, text: LocaleStrings.SelectaCustomDate }
];

export interface IChampionReportProps {
    context: WebPartContext;
    siteUrl: string;
    appTitle: string;
    loggedinUserEmail: string;
    onClickCancel: () => void;
}

export interface IChampionReportState {
    championsList: any;
    allChampionEvents: any;
    filteredAllEvents: any;
    selectedChampion: any;
    totalActivities: number;
    totalPoints: number;
    totalChampions: number;
    championRank: number;
    championsRank: any;
    noEventsFlag: boolean;
    eventTypesList: any;
    filterStartDate: Date;
    filterEndDate: Date;
    disableStartDate: boolean;
    disableEndDate: boolean;
    disableApplyButton: boolean;
    selectedFilterChoice: string;
    membersList: any;
    selectedMemberDetails: any;
    selectedDateFilter: any;
    isFilterCalloutVisible: boolean;
    showChampionEvents: boolean;
    eventsListHeight: string;
}

export default class ChampionReport extends React.Component<IChampionReportProps, IChampionReportState> {
    constructor(props: IChampionReportProps) {
        super(props);

        //States
        this.state = {
            championsList: "",
            allChampionEvents: "",
            filteredAllEvents: "",
            selectedChampion: "",
            totalActivities: 0,
            totalPoints: 0,
            totalChampions: 0,
            championRank: 0,
            championsRank: "",
            noEventsFlag: false,
            eventTypesList: [],
            filterStartDate: undefined,
            filterEndDate: undefined,
            disableStartDate: true,
            disableEndDate: true,
            disableApplyButton: true,
            selectedFilterChoice: stringsConstants.AllTime,
            membersList: "",
            selectedMemberDetails: [],
            selectedDateFilter: stringsConstants.AllTime,
            isFilterCalloutVisible: false,
            showChampionEvents: false,
            eventsListHeight: ""
        };

        //Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );

        this.updateEventsListHeight = this.updateEventsListHeight.bind(this);
    }

    //This method will be called when the component is mounted
    public async componentDidMount() {
        //get list of approved champions from Member list 
        await this.getAllChampions();

    }


    //Refresh the data in the report whenever the champion or datefilter is selected
    public componentDidUpdate(prevProps: Readonly<IChampionReportProps>, prevState: Readonly<IChampionReportState>, snapshot?: any): void {

        if (prevState.selectedChampion !== this.state.selectedChampion ||
            prevState.filteredAllEvents !== this.state.filteredAllEvents) {

            //load the data for the header cards
            this.getChampionMetrics(this.state.selectedChampion);

            //load the list of event types
            this.getEventsList(this.state.selectedChampion);

            if (this.state.selectedChampion !== stringsConstants.AllLabel && this.state.selectedChampion !== "") {
                this.setState({ showChampionEvents: true });
            }
            else {
                this.setState({ showChampionEvents: false });
            }
        }
    }

    //get list of approved champions from Member list and binding it to dropdown
    private async getAllChampions() {
        try {
            let filterQuery = "Status eq '" + constants.approvedStatus + "'";

            let allChampions: any[] = await commonServiceManager.getItemsWithOnlyFilter(
                stringsConstants.MemberList, filterQuery);
            let memberDetails: any = [];

            if (allChampions.length > 0) {
                //Sort on FirstName
                allChampions.sort((a, b) => a.FirstName.localeCompare(b.FirstName));
                let allChampionsChoices: IComboBoxOption[] = [{ key: stringsConstants.AllLabel, text: stringsConstants.AllChampionsLabel }];

                //Loop through all champions and create an array with key and text
                await allChampions.forEach((eachChampion) => {
                    allChampionsChoices.push({
                        key: eachChampion[stringsConstants.IDColumn],
                        text: eachChampion[stringsConstants.FirstName] + " " + eachChampion[stringsConstants.LastName],
                    });

                    memberDetails.push({
                        ID: eachChampion[stringsConstants.IDColumn],
                        Title: eachChampion[stringsConstants.TitleColumn],
                        FirstName: eachChampion[stringsConstants.FirstName],
                        LastName: eachChampion[stringsConstants.LastName]
                    });
                });

                this.setState({
                    championsList: allChampionsChoices,
                    selectedChampion: stringsConstants.AllLabel,
                    membersList: memberDetails
                });

                //get all champion events from Event Track Details List
                await this.getChampionEvents();
            }
        }
        catch (error) {
            console.error("CMP_ChampionReport_getAllChampions \n", error);
        }
    }

    //Get champion events from event track details list.
    private async getChampionEvents() {
        try {
            let allChampionEventsArray: any = [];
            let championRankArray: any = [];
            let filteredChampionEvents: any = [];
            let filterApprovedEvents = "Status eq 'Approved' or Status eq null or Status eq ''";

            //Get first batch of items from event track details list
            let championEventsArray = await commonServiceManager.getAllListItemsPagedWithFilter(constants.EventTrackDetailsList, filterApprovedEvents);
            if (championEventsArray.results.length > 0) {
                allChampionEventsArray.push(...championEventsArray.results);
                //Get next batch, if more items found in event track details list
                while (championEventsArray.hasNext) {
                    championEventsArray = await championEventsArray.getNext();
                    allChampionEventsArray.push(...championEventsArray.results);
                }
            }

            if (allChampionEventsArray.length > 0) {
                for (let i = 0; i < this.state.membersList.length; i++) {
                    filteredChampionEvents = allChampionEventsArray.filter((user: any) => user.MemberId === this.state.membersList[i].ID);

                    if (filteredChampionEvents.length > 0) {
                        //Sum up the points for each champion                      
                        let pointsCompleted: number = filteredChampionEvents.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue["Count"]; }, 0);

                        //Push the metrics of each champion into an array.
                        championRankArray.push({
                            Champion: this.state.membersList[i].FirstName,
                            Points: pointsCompleted,
                            MemberId: this.state.membersList[i].ID
                        });
                    }
                    else {
                        //Push the metrics of each champion into an array.
                        championRankArray.push({
                            Champion: this.state.membersList[i].FirstName,
                            Points: 0,
                            MemberId: this.state.membersList[i].ID
                        });
                    }
                }
                //Sort by points and then by Champion First Name
                championRankArray.sort((a: any, b: any) => {
                    if (a.Points < b.Points) return 1;
                    if (a.Points > b.Points) return -1;
                    if (a.Champion > b.Champion) return 1;
                    if (a.Champion < b.Champion) return -1;
                });

                this.setState({
                    allChampionEvents: allChampionEventsArray,
                    filteredAllEvents: allChampionEventsArray,
                    championsRank: championRankArray
                });
            }
            //If no events are found in the Event Track Details list, set the flag
            else this.setState({ noEventsFlag: true });
        }
        catch (error) {
            console.error("CMP_ChampionReport_getChampionEvents \n", JSON.stringify(error));
        }
    }

    //Get data for header cards 
    private async getChampionMetrics(selectedChampion: string) {
        try {
            this.setState({
                totalActivities: 0,
                totalPoints: 0,
                championRank: 0,
                totalChampions: 0,
                selectedMemberDetails: []
            });

            if (selectedChampion !== stringsConstants.AllLabel && selectedChampion !== "") {
                if (this.state.filteredAllEvents.length > 0) {
                    let championData = this.state.filteredAllEvents.filter((item: any) => item.MemberId === selectedChampion);
                    if (championData.length > 0) {
                        let totalEventPoints = championData.reduce(
                            (previousValue: any, currentValue: any) => { return previousValue + currentValue[stringsConstants.CountColumn]; }, 0);
                        let totalActivities = championData.length;
                        let index: number;
                        if (this.state.championsRank.findIndex((item: any) => item.MemberId === selectedChampion) !== -1) {
                            index = this.state.championsRank.findIndex((item: any) => item.MemberId === selectedChampion)
                        }
                        this.setState({
                            totalActivities: totalActivities,
                            totalPoints: totalEventPoints,
                            championRank: index + 1,
                            totalChampions: this.state.membersList.length
                        });
                    }
                    else {
                        let index: number;
                        if (this.state.championsRank.findIndex((item: any) => item.MemberId === selectedChampion) !== -1) {
                            index = this.state.championsRank.findIndex((item: any) => item.MemberId === selectedChampion)
                        }
                        this.setState({
                            totalActivities: 0,
                            totalPoints: 0,
                            championRank: index + 1,
                            totalChampions: this.state.membersList.length
                        });
                    }
                }

                //filter the selected champion from members list array
                let selectedMemberDetails = this.state.membersList.filter((item: any) => item.ID === selectedChampion);

                this.setState({
                    selectedMemberDetails: selectedMemberDetails
                })
            }
            else {

                if (this.state.filteredAllEvents.length > 0) {
                    //Calculating metrics for all champions
                    let totalEventPoints = this.state.filteredAllEvents.reduce(
                        (previousValue: any, currentValue: any) => { return previousValue + currentValue[stringsConstants.CountColumn]; }, 0);
                    let totalActivities = this.state.filteredAllEvents.length;

                    this.setState({
                        totalActivities: totalActivities,
                        totalPoints: totalEventPoints,
                        totalChampions: this.state.membersList.length
                    });
                }
            }
        }
        catch (error) {
            console.error("CMP_ChampionReport_getChampionMetrics \n", error);
        }
    }

    //get the list of events based on selected champion
    private async getEventsList(selectedChampion: string) {
        try {
            this.setState({
                eventTypesList: []
            });
            if (selectedChampion === stringsConstants.AllLabel) {
                if (this.state.filteredAllEvents.length > 0) {
                    let organizedEvents = commonServiceManager.groupBy(this.state.filteredAllEvents, (item: any) => item.EventName);
                    let topEventsArray: any[] = [];
                    // count the number of events for each event type
                    organizedEvents.forEach((event) => {
                        let totalEvents: number = event.length;

                        //Push the metrics of each event into an array.
                        topEventsArray.push({
                            Title: event[0].EventName,
                            Count: totalEvents
                        });
                    });

                    //Sort by number of events
                    topEventsArray.sort((a, b) => {
                        if (a.Count < b.Count) return 1;
                        if (a.Count > b.Count) return -1;
                    });

                    this.setState({
                        eventTypesList: topEventsArray
                    });
                }
            } else {
                if (this.state.filteredAllEvents.length > 0) {
                    let championEvents = this.state.filteredAllEvents.filter((item: any) => item.MemberId === selectedChampion);
                    if (championEvents.length > 0) {
                        let topEventsArray: any[] = [];
                        let organizedEvents = commonServiceManager.groupBy(championEvents, (item: any) => item.EventName);

                        // count the number of events for each event type
                        organizedEvents.forEach((event) => {
                            let totalEvents: number = event.length;

                            //Push the metrics of each event into an array.
                            topEventsArray.push({
                                Title: event[0].EventName,
                                Count: totalEvents
                            });
                        });

                        //Sort by number of events
                        topEventsArray.sort((a, b) => {
                            if (a.Count < b.Count) return 1;
                            if (a.Count > b.Count) return -1;
                        });
                        this.setState({
                            eventTypesList: topEventsArray
                        });
                    }
                }
            }
        }
        catch (error) {
            console.error("CMP_ChampionReport_getEventsList \n", error);
        }
    }

    //Set state variable when an option is selected in champion dropdown
    private setSelectedChampion = (ev: React.FormEvent<IComboBox>, option: IComboBoxOption): void => {
        this.setState({
            selectedChampion: option.key,
            eventTypesList: [],
            showChampionEvents: false
        });
    }

    //on select of date filter
    private onChoiceSelection = async (ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): Promise<void> => {

        let filteredEvents: any = "";

        switch (option.key) {
            case stringsConstants.AllTime:
                filteredEvents = this.state.allChampionEvents;
                break;

            case stringsConstants.LastMonth:
                filteredEvents = this.state.allChampionEvents.filter((o: any) => moment(new Date(o.DateofEvent), 'YYYY-MM-DD').isBetween(moment().subtract(1, 'months'), moment()));
                break;

            case stringsConstants.Last3Months:
                filteredEvents = this.state.allChampionEvents.filter((o: any) => moment(new Date(o.DateofEvent), 'YYYY-MM-DD').isBetween(moment().subtract(3, 'months'), moment()));
                break;

            case stringsConstants.Last6Months:
                filteredEvents = this.state.allChampionEvents.filter((o: any) => moment(new Date(o.DateofEvent), 'YYYY-MM-DD').isBetween(moment().subtract(6, 'months'), moment()));
                break;

            case stringsConstants.Last1Year:
                filteredEvents = this.state.allChampionEvents.filter((o: any) => moment(new Date(o.DateofEvent), 'YYYY-MM-DD').isBetween(moment().subtract(12, 'months'), moment()));
                break;

            default:
                break;
        }
        if (option.key === stringsConstants.SelectaCustomDate) {
            this.setState({
                disableStartDate: false,
                disableApplyButton: true,
                selectedFilterChoice: option.key,
                selectedDateFilter: option.key,
            });
        }
        else {
            this.setState({
                filterStartDate: undefined,
                filterEndDate: undefined,
                disableStartDate: true,
                disableEndDate: true,
                disableApplyButton: true,
                selectedDateFilter: option.key,
                selectedFilterChoice: option.key,
                filteredAllEvents: filteredEvents,
                isFilterCalloutVisible: false
            });
        }
    }

    //on select of Start Date in date picker
    private onStartDateChange(startDate: any) {
        this.setState({ filterStartDate: startDate });
        if (startDate !== undefined) {
            this.setState({
                disableEndDate: false
            });
        }
        if (this.state.filterEndDate !== undefined) {
            this.setState({
                disableApplyButton: false
            });
        }
    }

    //on select of End Date in date picker
    private onEndDateChange(endDate: any) {
        this.setState({ filterEndDate: endDate });
        this.setState({
            disableApplyButton: false
        });
    }

    //on click on apply button for custom date filter
    private onAppplyFilter() {
        let filteredEvents: any = "";
        let filteredChoice: string = ""
        if (this.state.selectedFilterChoice === stringsConstants.SelectaCustomDate) {

            filteredEvents = this.state.allChampionEvents.filter((item: any) =>
                new Date(item.DateofEvent) >= this.state.filterStartDate && new Date(item.DateofEvent) <= this.state.filterEndDate);
            filteredChoice = moment(this.state.filterStartDate).format("ddd MMM DD, YYYY") + " To " + moment(this.state.filterEndDate).format("ddd MMM DD, YYYY")
        }

        this.setState({
            filteredAllEvents: filteredEvents,
            selectedDateFilter: filteredChoice,
            isFilterCalloutVisible: false
        });
    }

    //on click on clear button in date filter
    private onClearFilter() {
        this.setState({
            filteredAllEvents: this.state.allChampionEvents,
            filterStartDate: undefined,
            filterEndDate: undefined,
            selectedFilterChoice: stringsConstants.AllTime,
            selectedDateFilter: stringsConstants.AllTime,
            disableApplyButton: true,
            disableStartDate: true,
            disableEndDate: true,
            isFilterCalloutVisible: false
        });
    }

    //Method to update Events list height in Pixels
    private updateEventsListHeight(height: string) {
        this.setState({ eventsListHeight: height });
    }


    //Render method
    public render() {
        // added constants for filter dropdown menu
        const filterButtonId = 'date-filter-callout-button';
        const filterLabelId = 'date-filter-callout-label';
        const filterDescriptionId = 'date-filter-callout-description';
        return (
            <>
                <div className={styles.container}>
                    <div className={styles.cmpReportPath}>
                        <img src={require("../assets/TOTImages/BackIcon.png")}
                            className={styles.backImg}
                            alt={LocaleStrings.BackButton}
                            aria-hidden="true"
                        />
                        <span
                            className={styles.backLabel}
                            onClick={() => this.props.onClickCancel()}
                            title={this.props.appTitle}
                        >
                            {this.props.appTitle}
                        </span>
                        <span className={styles.border}></span>
                        <span className={styles.cmpReportLabel}>{LocaleStrings.ChampionReportLabel}</span>
                    </div>
                    <br />
                    {this.state.noEventsFlag ?
                        <div>{LocaleStrings.NoEventsMessage}</div>
                        :
                        <>
                            {this.state.championsList.length > 0 && (
                                <Row xl={2} lg={2} md={2} sm={1} className="selection-filter-area">
                                    <Col xl={5} lg={6} md={6} sm={12}>
                                        <div className={styles.cmpReportDropdownWrapper}>
                                            <ComboBox
                                                label={LocaleStrings.ChampionLabel}
                                                defaultSelectedKey={stringsConstants.AllLabel}
                                                options={this.state.championsList}
                                                onChange={this.setSelectedChampion.bind(this)}
                                                className={styles.cmpReportDropdown}
                                                calloutProps={{ className: "cmpReportComboBoxCallout" }}
                                            />
                                        </div>
                                    </Col>
                                    <Col xl={7} lg={6} md={6} sm={12}>
                                        <div className={styles.dateFilterCtrl}>
                                            <label className={styles.dateFilterLabel}> {LocaleStrings.FilterbyDateLabel} </label>
                                            <div
                                                className={`${styles.dateFilterDropdown}${this.state.isFilterCalloutVisible ? " " + styles.calloutVisible : ""}`}
                                                onClick={() => this.setState({ isFilterCalloutVisible: !this.state.isFilterCalloutVisible })}
                                                id={filterButtonId}
                                            >
                                                <span className={styles.filterLabel} title={this.state.selectedDateFilter}>
                                                    {this.state.selectedDateFilter}
                                                </span>
                                                {this.state.isFilterCalloutVisible ?
                                                    <Icon iconName="ChevronUp" />
                                                    :
                                                    <Icon iconName="ChevronDown" />
                                                }
                                            </div>
                                            {this.state.isFilterCalloutVisible && (
                                                <Callout
                                                    ariaLabelledBy={filterLabelId}
                                                    ariaDescribedBy={filterDescriptionId}
                                                    role="dialog"
                                                    className="filter-links-callout"
                                                    gapSpace={0}
                                                    target={`#${filterButtonId}`}
                                                    isBeakVisible={false}
                                                    directionalHint={DirectionalHint.bottomLeftEdge}
                                                    onDismiss={() => this.setState({ isFilterCalloutVisible: false })}
                                                    preventDismissOnEvent={(eve) => {
                                                        if (eve.type === "scroll") return true;
                                                        else return false
                                                    }}
                                                    doNotLayer={true}
                                                >
                                                    <ChoiceGroup
                                                        options={options}
                                                        onChange={this.onChoiceSelection.bind(this)}
                                                        selectedKey={this.state.selectedFilterChoice}
                                                        defaultSelectedKey={stringsConstants.AllTime}
                                                        className="date-filter-choice-grp"
                                                    />
                                                    <div className="date-picker-wrapper">
                                                        <DatePicker
                                                            label={LocaleStrings.FromLabel}
                                                            allowTextInput
                                                            placeholder={LocaleStrings.SelectDate}
                                                            ariaLabel={LocaleStrings.SelectDate}
                                                            disabled={this.state.disableStartDate}
                                                            value={this.state.filterStartDate}
                                                            onSelectDate={this.onStartDateChange.bind(this)}
                                                            strings={defaultDatePickerStrings}
                                                            calloutProps={{ className: "reportDatePickerCallout" }}
                                                            calendarProps={{ className: "reportCalendarProps", strings: null }}
                                                        />
                                                        <DatePicker
                                                            label={LocaleStrings.ToLabel}
                                                            allowTextInput
                                                            placeholder={LocaleStrings.SelectDate}
                                                            ariaLabel={LocaleStrings.SelectDate}
                                                            value={this.state.filterEndDate}
                                                            disabled={this.state.disableEndDate}
                                                            minDate={this.state.filterStartDate}
                                                            onSelectDate={this.onEndDateChange.bind(this)}
                                                            strings={defaultDatePickerStrings}
                                                            calloutProps={{ className: "reportDatePickerCallout" }}
                                                            calendarProps={{ className: "reportCalendarProps", strings: null }}
                                                        />
                                                    </div>
                                                    <div className='filter-btn-wrapper'>
                                                        <PrimaryButton
                                                            onClick={this.onAppplyFilter.bind(this)}
                                                            text={LocaleStrings.ApplyButton}
                                                            title={LocaleStrings.ApplyButton}
                                                            disabled={this.state.disableApplyButton}
                                                        />
                                                        <DefaultButton
                                                            onClick={this.onClearFilter.bind(this)}
                                                            title={LocaleStrings.ClearButton}
                                                            text={LocaleStrings.ClearButton}
                                                        />
                                                    </div>
                                                </Callout>
                                            )}
                                        </div>
                                    </Col>
                                </Row>
                            )}
                            {this.state.selectedChampion &&
                                <Row xl={3} lg={3} md={2} sm={2} xs={1}>
                                    <Col xl={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4}
                                        lg={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalActivitiesLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalActivities}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                    <Col xl={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4} lg={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalPointsLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalPoints}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                    {this.state.selectedChampion !== stringsConstants.AllLabel &&
                                        <Col xl={3} lg={3} md={6} sm={6} xs={12}>
                                            <Card className={styles.mainCard}>
                                                <Card.Body className={styles.cardBody}>
                                                    <Card.Text className={styles.cardTitleText}>
                                                        {LocaleStrings.RankLabel}
                                                    </Card.Text>
                                                    <Card.Title className={styles.cardValue}>{this.state.championRank}</Card.Title>
                                                </Card.Body>
                                            </Card>
                                        </Col>
                                    }
                                    <Col xl={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4} lg={this.state.selectedChampion !== stringsConstants.AllLabel ? 3 : 4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalChampionsLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalChampions}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                </Row>
                            }

                            {this.state.selectedChampion &&
                                <Row xl={2} lg={2} md={2} sm={1} xs={1} className="events-chart-wrapper">
                                    <Col xl={4} lg={4} md={6} sm={12} xs={12}>
                                        <div className={styles.eventTypesWrapper}>
                                            <div className={styles.eventTypesHeading}>{LocaleStrings.EventTypesHeading}</div>
                                            <div
                                                className={styles.eventTypesList}
                                                style={{ minHeight: this.state.eventsListHeight, maxHeight: this.state.eventsListHeight }}
                                            >
                                                {this.state.eventTypesList.length === 0 ?
                                                    <div className={styles.eventItem + " " + styles.noEventsMessage}>
                                                        {LocaleStrings.NoEventsFoundMessage}
                                                    </div>
                                                    :
                                                    <>
                                                        {this.state.eventTypesList.map((element: any) =>
                                                            (<div className={styles.eventItem}> {element.Title} </div>)
                                                        )}
                                                    </>
                                                }
                                            </div>
                                        </div>
                                    </Col>
                                    <Col xl={8} lg={8} md={6} sm={12} xs={12}>
                                        <EventsChart
                                            siteUrl={this.props.siteUrl}
                                            context={this.props.context}
                                            filteredAllEvents={this.state.filteredAllEvents}
                                            parentComponent={stringsConstants.ChampionReportLabel}
                                            selectedMemberID={this.state.selectedChampion}
                                            updateEventsListHeight={this.updateEventsListHeight}
                                        />
                                    </Col>
                                </Row>
                            }

                            {this.state.showChampionEvents &&
                                <ChampionEvents
                                    context={this.props.context}
                                    filteredAllEvents={this.state.filteredAllEvents}
                                    selectedMemberDetails={this.state.selectedMemberDetails}
                                    parentComponent={stringsConstants.ChampionReportLabel}
                                    selectedMemberID={this.state.selectedChampion}
                                    loggedinUserEmail={this.props.loggedinUserEmail}
                                />
                            }
                        </>
                    }
                </div>
            </>
        );
    }
}


