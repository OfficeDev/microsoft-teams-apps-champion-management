import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as stringsConstants from '../constants/strings';
import BootstrapTable from 'react-bootstrap-table-next';
import Card from 'react-bootstrap/esm/Card';
import Col from 'react-bootstrap/esm/Col';
import commonServices from '../Common/CommonServices';
import paginationFactory from 'react-bootstrap-table2-paginator';
import React from 'react';
import Row from 'react-bootstrap/esm/Row';
import styles from '../scss/TOTReport.module.scss';
import ToolkitProvider, { CSVExport, Search, ToolkitContextType } from 'react-bootstrap-table2-toolkit';
import { Chart } from 'chart.js';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import {
    ComboBox,
    DirectionalHint,
    IComboBox,
    IComboBoxOption,
    Icon,
    Spinner,
    SpinnerSize,
    TooltipHost
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';


//Global Variables
let commonServiceManager: commonServices;
const chartLabelLength: number = 14;
const horizontalChartColor: any = ['#02A5F2', '#FFBA02', '#888686', '#5BBB02', '#FF4F17'];
const doughnutChartColor: any = ['#012357', '#02A5F2'];

export interface ITOTReportProps {
    context?: WebPartContext;
    siteUrl: string;
    onClickCancel: () => void;
}

export interface ITOTReportState {
    searchedString: string;
    tournamentsList: any;
    noTournamentsFlag: boolean;
    selectedTournament: any;
    totalActivities: number;
    totalPoints: number;
    totalParticipants: number;
    totalCompletionPercentage: number;
    participantsList: any;
    topParticipantsChartdata: Chart.ChartData;
    topTournamentsChartdata: Chart.ChartData;
    participantsStatusChartdata: Chart.ChartData;
    csvFileName: string;
    showSpinner: boolean;
}

export default class TOTReport extends React.Component<ITOTReportProps, ITOTReportState> {
    constructor(props: ITOTReportProps, state: ITOTReportState) {
        super(props);
        this.state = {
            searchedString: "",
            tournamentsList: "",
            noTournamentsFlag: false,
            selectedTournament: "",
            totalActivities: 0,
            totalPoints: 0,
            totalParticipants: 0,
            totalCompletionPercentage: 0,
            participantsList: [],
            topParticipantsChartdata: {},
            topTournamentsChartdata: {},
            participantsStatusChartdata: {},
            csvFileName: "",
            showSpinner: false

        };

        //Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );
        this.progressBar = this.progressBar.bind(this);
        this.loadParticipantsStatus = this.loadParticipantsStatus.bind(this);
        this.loadTopParticipants = this.loadTopParticipants.bind(this);
        this.loadTopTournaments = this.loadTopTournaments.bind(this);
        this.clearSearch = this.clearSearch.bind(this);
    }

    //On Load method
    public componentDidMount() {
        //Get list of completed tournaments from Tournaments list
        this.getCompletedTournaments();
    }

    //Refresh the data in the report whenever the tournament name is selected
    public componentDidUpdate(prevProps: Readonly<ITOTReportProps>, prevState: Readonly<ITOTReportState>, snapshot?: any): void {
        if (prevState.selectedTournament != this.state.selectedTournament) {

            //Load the data for the header cards
            this.getTournamentMetrics(this.state.selectedTournament);
            //Load participants data into grid table    
            this.getParticipantsList(this.state.selectedTournament);
            //Load the Top 5 Participants Chart
            this.loadTopParticipants();
            //Load the Top 5 Tournaments Chart
            this.loadTopTournaments();
            //Load the Participants Status Chart
            this.loadParticipantsStatus();

        }
    }

    //Get list of completed tournaments from Tournaments Report list and binding it to dropdown
    private async getCompletedTournaments() {
        try {
            let completedTournamentDetails: any[] = await commonServiceManager.getAllListItems(
                stringsConstants.TournamentsReportList);

            if (completedTournamentDetails.length > 0) {
                //Sort on Title
                completedTournamentDetails.sort((a, b) => a.Title.localeCompare(b.Title));
                let completedTournamentsChoices: IComboBoxOption[] = [{ key: stringsConstants.AllLabel, text: stringsConstants.AllTournamentsLabel }];

                //Loop through all "Completed" tournaments and create an array with key and text
                await completedTournamentDetails.forEach((eachTournament) => {
                    if (eachTournament[stringsConstants.CompletedOnColumn]) {
                        let completedDate = commonServiceManager.formatDate(new Date(eachTournament[stringsConstants.CompletedOnColumn]));
                        completedTournamentsChoices.push({
                            key: eachTournament[stringsConstants.TitleColumn],
                            text: eachTournament[stringsConstants.TitleColumn] + " - " + completedDate,
                        });
                    } else {
                        completedTournamentsChoices.push({
                            key: eachTournament[stringsConstants.TitleColumn],
                            text: eachTournament[stringsConstants.TitleColumn],
                        });
                    }
                });

                this.setState({
                    tournamentsList: completedTournamentsChoices,
                    selectedTournament: stringsConstants.AllLabel,
                    csvFileName: stringsConstants.AllLabel
                });
            }
            //If no completed tournaments are found in the Tournaments Report list, set the flag
            else this.setState({ noTournamentsFlag: true });
        }
        catch (error) {
            console.error("TOT_TOTReport_getCompletedTournaments \n", error);
        }
    }

    //Get data for header cards from Tournaments Report List
    private async getTournamentMetrics(completedTournament: string) {
        try {
            if (completedTournament != stringsConstants.AllLabel) {
                let filter: string = "Title eq '" + completedTournament.trim().replace(/'/g, "''") + "'";
                const tournamentData: any[] = await commonServiceManager.getItemsWithOnlyFilter(
                    stringsConstants.TournamentsReportList, filter);
                if (tournamentData.length > 0) {
                    this.setState({
                        totalActivities: tournamentData[0].Total_x0020_Activities,
                        totalPoints: tournamentData[0].Total_x0020_Points,
                        totalParticipants: tournamentData[0].Total_x0020_Participants,
                        totalCompletionPercentage: tournamentData[0].Completion_x0020_Percentage
                    });
                }
            }
            else {
                const tournamentData: any[] = await commonServiceManager.getAllListItems(
                    stringsConstants.TournamentsReportList);
                if (tournamentData.length > 0) {
                    //Calculating metrics for all tournaments
                    let totalTournamentActivities = tournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.TotalActivitiesColumn]; }, 0);
                    let totalTournamentPoints = tournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.TotalPointsColumn]; }, 0);
                    let totalTournamentParticipants = tournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.TotalParticipantsColumn]; }, 0);
                    let totalTournamentCompletedParticipants = tournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.CompletedParticipantsColumn]; }, 0);
                    let totalPercentage = Math.round(totalTournamentCompletedParticipants * 100 / totalTournamentParticipants);

                    this.setState({
                        totalActivities: totalTournamentActivities,
                        totalPoints: totalTournamentPoints,
                        totalParticipants: totalTournamentParticipants,
                        totalCompletionPercentage: totalPercentage
                    });
                }
            }
        }
        catch (error) {
            console.error("TOT_TOTReport_getTournamentMetrics \n", error);

        }
    }

    //Get data for Participants Grid for selected tournament from Participants Report list
    private async getParticipantsList(completedTournament: string) {
        try {
            if (completedTournament != stringsConstants.AllLabel) {
                let participantsData: any = [];
                let filter: string = "Title eq '" + this.state.selectedTournament.trim().replace(/'/g, "''") + "'";
                const allItems: any[] = await commonServiceManager.getItemsSortedWithFilter(
                    stringsConstants.ParticipantsReportList, filter, stringsConstants.PointsColumn);
                if (allItems.length > 0) {
                    for (let i = 0; i < allItems.length; i++) {
                        participantsData.push({
                            name: allItems[i].User_x0020_Name,
                            activitiesCompleted: allItems[i].Activities_x0020_Completed,
                            points: allItems[i].Points,
                            tournamentCompletedPercentage: allItems[i].Completion_x0020_Percentage
                        });
                    }
                    this.setState({
                        participantsList: participantsData
                    });
                }
                else {
                    this.setState({
                        participantsList: []
                    });
                }

            }
            else {
                this.setState({
                    participantsList: []
                });
            }
        }
        catch (error) {
            console.error("TOT_TOTReport_getParticipantsList \n", error);
        }
    }

    
    //Load Top Participants based on points
    private async loadTopParticipants(): Promise<Chart.ChartData> {
        try {

            let arrLabels: string[] = [];
            let arrData: number[] = [];
            let topParticipants: any[] = [];

            if (this.state.selectedTournament == stringsConstants.AllLabel) {
                let allParticipantsArray: any = [];
                let topParticipantsArray: any = [];

                this.setState({
                    showSpinner: true
                });

                //Get first batch of items from Participants Report list
                let participantsArray = await commonServiceManager.getAllListItemsPaged(stringsConstants.ParticipantsReportList);
                if (participantsArray.results.length > 0) {
                    allParticipantsArray.push(...participantsArray.results);
                    //Get next batch, if more items found in Participants Report list
                    while (participantsArray.hasNext) {
                        participantsArray = await participantsArray.getNext();
                        allParticipantsArray.push(...participantsArray.results);
                    }
                }
                //Group the items by participants
                let organizedParticipants = commonServiceManager.groupBy(allParticipantsArray, item => item.User_x0020_Name);

                //Sum up the points for each participant
                organizedParticipants.forEach((participant) => {
                    let pointsCompleted: number = participant.reduce((previousValue, currentValue) => { return previousValue + currentValue["Points"]; }, 0);

                    //Push the metrics of each participant into an array.
                    topParticipantsArray.push({
                        Title: participant[0].User_x0020_Name,
                        Points: pointsCompleted
                    });
                });

                //Sort by points and then by user name 
                topParticipantsArray.sort((a, b) => {
                    if (a.Points < b.Points) return 1;
                    if (a.Points > b.Points) return -1;
                    if (a.Title > b.Title) return 1;
                    if (a.Title < b.Title) return -1;
                });

                //Get Top 5 participants and add it to Top Participants List
                let top5ParticipantsArray = topParticipantsArray.filter((item, idx) => idx < 5).map(item => { return item; });

                if (top5ParticipantsArray.length > 0) {
                    //Loop through top tournaments and add the lables and data for the chart display
                    top5ParticipantsArray.forEach(element => {
                        arrLabels.push(element.Title);
                        arrData.push(element.Points);
                    });
                }
            }
            else {

                let columns = stringsConstants.UserNameColumn + "," + stringsConstants.PointsColumn;
                let descColumn = stringsConstants.PointsColumn;
                let ascColumn = stringsConstants.UserNameColumn;
                let filter: string = "Title eq '" + this.state.selectedTournament.trim().replace(/'/g, "''") + "'";

                topParticipants =
                    await commonServiceManager.getFilteredTopSortedItemsWithSpecificColumns(
                        stringsConstants.ParticipantsReportList, filter, columns, 5, descColumn, ascColumn
                    );
                if (topParticipants.length > 0) {
                    //Loop through top tournaments and add the lables and data for the chart display
                    topParticipants.forEach(element => {
                        arrLabels.push(element.User_x0020_Name);
                        arrData.push(element.Points);
                    });
                }
            }

            this.setState({
                topParticipantsChartdata: {
                    labels: arrLabels,
                    datasets: [{
                        label: '',
                        data: arrData,
                        backgroundColor: horizontalChartColor,
                        fill: false
                    }]
                },
                showSpinner: false,
            });
            //Resetting the state to render the chart data with current state value
            setTimeout(() => {
                this.setState({
                    topParticipantsChartdata: this.state.topParticipantsChartdata,
                });
            }, 10);

        }
        catch (error) {
            console.error("TOT_TOTReport_loadTopParticipants \n", error);
        }
    }

    //Load Top tournaments based on number of participants   
    private async loadTopTournaments(): Promise<Chart.ChartData> {
        try {
            let arrLabels: string[] = [];
            let arrData: number[] = [];
            let columns = stringsConstants.TitleColumn + "," + stringsConstants.TotalParticipantsColumn;
            let ascColumn = stringsConstants.CompletedOnColumn;
            let descColumn = stringsConstants.TotalParticipantsColumn;

            const topTournaments: any[] = await commonServiceManager.getTopSortedItemsWithSpecificColumns(stringsConstants.TournamentsReportList,
                columns, 5, descColumn, ascColumn);
            if (topTournaments.length > 0) {
                //Loop through top tournaments and add the lables and data for the chart display
                topTournaments.forEach(element => {
                    arrLabels.push(element.Title);
                    arrData.push(element.Total_x0020_Participants);
                });
            }
            this.setState({
                topTournamentsChartdata: {
                    labels: arrLabels,
                    datasets: [{
                        label: '',
                        data: arrData,
                        backgroundColor: horizontalChartColor,
                        fill: false
                    }]
                }
            });
            //Resetting the state to render the chart data with current state value
            setTimeout(() => {
                this.setState({
                    topTournamentsChartdata: this.state.topTournamentsChartdata,
                });
            }, 10);

        }
        catch (error) {
            console.error("TOT_TOTReport_loadTopTournaments \n", error);
        }

    }

    //Load particpants status chart based on selected tournament
    private async loadParticipantsStatus(): Promise<Chart.ChartData> {
        try {

            let arrLabels: string[] = [stringsConstants.CompletedChartLabel, stringsConstants.NotCompletedChartLabel];
            let arrData: number[] = [];
            let columns = stringsConstants.TotalParticipantsColumn + "," + stringsConstants.CompletedParticipantsColumn;

            if (this.state.selectedTournament == stringsConstants.AllLabel) {

                //Get the data for all tournaments from "Tournaments Report" list
                const allTournamentData: any[] = await commonServiceManager.getAllItemsWithSpecificColumns(
                    stringsConstants.TournamentsReportList, columns);

                if (allTournamentData.length > 0) {
                    let totalTournamentParticipants = allTournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.TotalParticipantsColumn]; }, 0);
                    let totalTournamentCompletedParticipants = allTournamentData.reduce(
                        (previousValue, currentValue) => { return previousValue + currentValue[stringsConstants.CompletedParticipantsColumn]; }, 0);

                    arrData.push(totalTournamentCompletedParticipants, totalTournamentParticipants - totalTournamentCompletedParticipants);
                }
            }
            else {

                let filter: string = "Title eq '" + this.state.selectedTournament.trim().replace(/'/g, "''") + "'";

                //Get the data for the selected tournament from "Tournaments Report" list
                const participantsStatusDetails: any[] = await commonServiceManager.getFilteredListItemsWithSpecificColumns(
                    stringsConstants.TournamentsReportList, columns, filter);

                //Pass the labels and data to the chart
                if (participantsStatusDetails.length > 0) {
                    arrData.push(participantsStatusDetails[0].Completed_x0020_Participants, participantsStatusDetails[0].Total_x0020_Participants - participantsStatusDetails[0].Completed_x0020_Participants);
                }
            }

            //Setting state
            this.setState({
                participantsStatusChartdata: {
                    labels: arrLabels,
                    datasets: [{
                        data: arrData,
                        backgroundColor: doughnutChartColor,
                        fill: false
                    }]
                }
            });
            //Resetting the state to render the chart data with current state value
            setTimeout(() => {
                this.setState({
                    participantsStatusChartdata: this.state.participantsStatusChartdata,
                });
            }, 10);


        }
        catch (error) {
            console.error("TOT_TOTReport_loadParticipantsStatus \n", error);
        }

    }

    //Options for horizontal bar chart
    private horizontalChartOptions = {
        legend: {
            labels: {
                boxWidth: 0,
            }
        }, scales: {
            xAxes: [{
                stacked: true,
                ticks: {
                    fontColor: "#020000",
                    fontSize: 14,
                    tooltips: true,
                    stepSize: 20,
                    callback: (label: string) => {
                        if (label.length > chartLabelLength) {
                            return label.slice(0, chartLabelLength) + '...';
                        } else {
                            return label;
                        }
                    }
                }
            }],
            yAxes: [{
                stacked: true,
                ticks: {
                    fontColor: "#020000",
                    fontSize: 14,
                    tooltips: true,
                    callback: (label: string) => {
                        if (label.length > chartLabelLength) {
                            return label.slice(0, chartLabelLength) + '...';
                        } else {
                            return label;
                        }
                    }
                }
            }]
        },
    };

    //Options for doughnut chart
    private doughNutChartOptions = {
        legend: {
            display: true,
            position: "bottom"
        }
    };

    //Set state variable when an option is selected Tournaments dropdown
    private setSelectedTournament = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
        this.setState({
            selectedTournament: option.key,
            csvFileName: option.text,
            topParticipantsChartdata: {},
            topTournamentsChartdata: {},
            participantsStatusChartdata: {}
        });
    }

    //Clear text in Search control
    private clearSearch = (onSearch: any) => {
        const searchElement = document.getElementById('search-bar-0') as HTMLInputElement;
        searchElement.value = '';
        onSearch('');
    }

    //Percentage Progressbar
    private progressBar = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        return (
            <div>
                <span className={styles.percentage}>{gridRow.tournamentCompletedPercentage}% </span>
                <progress value={gridRow.tournamentCompletedPercentage} max={100} className={styles.progressBar} />
            </div>
        );
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
        alwaysShowAllBtns: false
    });

    //Render method
    public render() {
        const { ExportCSVButton } = CSVExport;
        const participantsTableHeader = [
            {
                dataField: stringsConstants.Name,
                text: LocaleStrings.NameLabel,
                sort: true,
                headerTitle: true,
                title: true,
            }, {
                dataField: stringsConstants.ActivitiesCompleted,
                text: LocaleStrings.ActivitiesCompletedLabel,
                headerTitle: true,
                title: true
            }, {
                dataField: stringsConstants.Points,
                text: LocaleStrings.PointsLabel,
                sort: true,
                headerTitle: true,
                title: true
            }, {
                dataField: stringsConstants.TournamentCompletedPercentage,
                text: LocaleStrings.PercentageTournamentCompletedLabel,
                headerTitle: true,
                title: true,
                formatter: this.progressBar
            }];
        const { SearchBar } = Search;
        return (
            <>
                <div className={styles.container}>
                    <div className={styles.totReportPath}>
                        <img src={require("../assets/TOTImages/BackIcon.png")}
                            className={styles.backImg}
                            alt={LocaleStrings.BackButton}
                        />
                        <span
                            className={styles.backLabel}
                            onClick={() => this.props.onClickCancel()}
                            title={LocaleStrings.TOTBreadcrumbLabel}
                        >
                            {LocaleStrings.TOTBreadcrumbLabel}
                        </span>
                        <span className={styles.border}></span>
                        <span className={styles.totReportLabel}>{LocaleStrings.TournamentReportsPageTitle}</span>
                    </div>
                    <br />
                    {this.state.noTournamentsFlag ?
                        <div>{LocaleStrings.NoCompletedTournamentsMessage}</div>
                        :
                        <div>
                            {this.state.tournamentsList.length > 0 && (
                                <div>
                                    <ComboBox
                                        label={LocaleStrings.TournamentLabel}
                                        defaultSelectedKey={stringsConstants.AllLabel}
                                        options={this.state.tournamentsList}
                                        onChange={this.setSelectedTournament.bind(this)}
                                        className={styles.totReportDropdown}
                                        calloutProps={{ className: "totReportComboBoxCallout" }}
                                    />
                                    <span className={styles.dDInfoIconArea}>
                                        <TooltipHost
                                            content={LocaleStrings.ReportsDropdownInfoIconText}
                                            delay={2}
                                            directionalHint={DirectionalHint.topCenter}
                                        >
                                            <Icon iconName='Info' className={styles.dDInfoIcon} aria-label={LocaleStrings.ReportsDropdownInfoIconText} />
                                        </TooltipHost>
                                    </span>
                                </div>
                            )}
                            <br />

                            {this.state.selectedTournament &&
                                <Row xl={4} lg={3} md={2} sm={2} xs={1}>
                                    <Col xl={3} lg={4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalActivitiesLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalActivities}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                    <Col xl={3} lg={4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalPointsLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalPoints}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                    <Col xl={3} lg={4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalParticipantsLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalParticipants}</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                    <Col xl={3} lg={4} md={6} sm={6} xs={12}>
                                        <Card className={styles.mainCard}>
                                            <Card.Body className={styles.cardBody}>
                                                <Card.Text className={styles.cardTitleText}>
                                                    {LocaleStrings.TotalPercentageCompletionLabel}
                                                </Card.Text>
                                                <Card.Title className={styles.cardValue}>{this.state.totalCompletionPercentage}%</Card.Title>
                                            </Card.Body>
                                        </Card>
                                    </Col>
                                </Row>
                            }
                            <br />
                            {this.state.selectedTournament &&
                                <Row xl={3} lg={2} md={2} sm={1} xs={1}>
                                    <Col xl={4} lg={6} md={6} sm={12} xs={12}>
                                        <div className={styles.chartArea}>
                                            <span className={styles.chartHeading}>{LocaleStrings.Top5ParticipantswithPointsLabel}</span>
                                            <span>
                                                {this.state.showSpinner && (
                                                    <Spinner
                                                        size={SpinnerSize.large}
                                                        ariaLabel={LocaleStrings.LoadingSpinnerLabel}
                                                        label={LocaleStrings.LoadingSpinnerLabel}
                                                        ariaLive="assertive"
                                                    />
                                                )}
                                            </span>
                                                <ChartControl
                                                    type={ChartType.HorizontalBar}
                                                    data={this.state.topParticipantsChartdata}
                                                    options={this.horizontalChartOptions}
                                                    accessibility={{ enable: true, alternateText: `${LocaleStrings.Top5ParticipantswithPointsLabel} chart` }}
                                                />
                                            
                                        </div>
                                    </Col>
                                    <Col xl={4} lg={6} md={6} sm={12} xs={12}>
                                        <div className={styles.chartArea}>
                                            <span className={styles.chartHeading}>{LocaleStrings.Top5TournamentswithParticipantsLabel}</span>
                                            <ChartControl
                                                type={ChartType.HorizontalBar}
                                                data={this.state.topTournamentsChartdata}
                                                options={this.horizontalChartOptions}
                                                accessibility={{ enable: true, alternateText: `${LocaleStrings.Top5TournamentswithParticipantsLabel} chart` }}
                                            />
                                        </div>
                                    </Col>
                                    <Col xl={4} lg={6} md={6} sm={12} xs={12}>
                                        <div className={styles.chartArea}>
                                            <span className={styles.chartHeading}>{LocaleStrings.ParticipantsStatusLabel}</span>
                                            <ChartControl
                                                type={ChartType.Doughnut}
                                                data={this.state.participantsStatusChartdata}
                                                options={this.doughNutChartOptions}
                                                accessibility={{ enable: true, alternateText: `${LocaleStrings.ParticipantsStatusLabel} chart` }}
                                            />
                                        </div>
                                    </Col>
                                </Row>
                            }
                            <br />
                            {this.state.participantsList.length > 0 && (
                                <ToolkitProvider
                                    bootstrap4
                                    keyField={stringsConstants.Name}
                                    data={this.state.participantsList}
                                    columns={participantsTableHeader}
                                    exportCSV={{ fileName: `${this.state.csvFileName}.csv` }}
                                    search
                                >
                                    {
                                        (props: ToolkitContextType) => (
                                            <div>
                                                <div className={styles.searchArea}>
                                                    <SearchBar {...props.searchProps} placeholder={LocaleStrings.SearchPlaceholder} />
                                                    <div className={styles.IconArea}>
                                                        {props.searchProps.searchText ?
                                                            <Icon
                                                                iconName='Cancel'
                                                                onClick={() => this.clearSearch(props.searchProps.onSearch)}
                                                                className={styles.cancelIcon}
                                                            />
                                                            :
                                                            <Icon
                                                                iconName='Search'
                                                                className={styles.searchIcon}
                                                            />
                                                        }
                                                    </div>

                                                </div>
                                                <div className={styles.tableHeadingAndCSVBtn}>
                                                    <div className={styles.tableHeading}>{LocaleStrings.ParticipantsDetailsLabel}</div>
                                                    <ExportCSVButton {...props.csvProps} className={styles.csvLink}>
                                                        <img
                                                            src={require("../assets/TOTImages/ExcelIcon.png")}
                                                            alt="Export Icon"
                                                            className={styles.excelIcon}
                                                        />
                                                        <span className={styles.downloadText}>{LocaleStrings.DownloadButtonText}</span>
                                                    </ExportCSVButton>
                                                </div>
                                                <div className={styles.participantsGrid}>
                                                    <BootstrapTable
                                                        {...props.baseProps}
                                                        table-responsive={true}
                                                        defaultSorted={[{ dataField: stringsConstants.Points, order: 'desc' }]}
                                                        pagination={this.pagination}
                                                        noDataIndication={() => (<div>{LocaleStrings.NoRecordsinGridLabel}</div>)}
                                                    />
                                                </div>
                                            </div>
                                        )
                                    }
                                </ToolkitProvider>
                            )}
                        </div>
                    }
                </div>
            </>
        );
    }
}


