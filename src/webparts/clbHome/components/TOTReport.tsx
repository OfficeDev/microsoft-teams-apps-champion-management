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
    TooltipHost,
    Label
} from '@fluentui/react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";


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
    private totReportGridRef: React.RefObject<HTMLDivElement>;
    private totReportComboboxWrapperRef: React.RefObject<HTMLDivElement>;
    private mainComboboxRef: React.RefObject<IComboBox>;
    constructor(props: ITOTReportProps, state: ITOTReportState) {
        super(props);
        this.totReportGridRef = React.createRef();
        this.totReportComboboxWrapperRef = React.createRef();
        this.mainComboboxRef = React.createRef();
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
        try {
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
            //Accessibility
            //Update Page Per Size Dropdown and its button to be accessibile through Keyboard
            if (prevState.participantsList !== this.state.participantsList && this.state.participantsList.length > 0) {
                //Outside Click event for page dropdown button
                document.addEventListener("click", this.onPageDropdownBtnOutsideClick);

                //Set aria-label to participant details table element
                const tableElement = this.totReportGridRef?.current?.querySelector("table");
                tableElement.setAttribute("aria-label", LocaleStrings.ParticipantsDetailsLabel);

                //Get Page Dropdown button element
                const sizePerPageBtnElement: any = this.totReportGridRef?.current?.querySelector("#pageDropDown");
                sizePerPageBtnElement.setAttribute("aria-label", stringsConstants.sizePerPageLabel + " " + sizePerPageBtnElement?.textContent);

                //Update Page Dropdown button click event
                sizePerPageBtnElement.addEventListener("click", (evt: any) => {

                    const sizePerPageUlElement = this.getUlElement();
                    if (sizePerPageUlElement.getAttribute("style") === "display:block") {
                        sizePerPageUlElement.setAttribute("style", "display:none");
                    }
                    else {
                        sizePerPageUlElement.setAttribute("style", "display:block");
                    }
                });

                //Get Page Dropdown Callout element
                const sizePerPageUlElement = this.getUlElement();

                //Get Page Size anchor Elements
                const pageSizeAnchorElements: any = sizePerPageUlElement.getElementsByTagName("a");

                //Update all page size option elements to support access with keyboard arrow keys
                for (let i = 0; i < pageSizeAnchorElements?.length; i++) {
                    pageSizeAnchorElements[i]?.addEventListener("keydown", (event: any) => {
                        if (event.keyCode === 38 && i > 0) {
                            pageSizeAnchorElements[i - 1]?.focus();
                        }
                        else if (event.keyCode === 40 && i < pageSizeAnchorElements?.length - 1) {
                            pageSizeAnchorElements[i + 1]?.focus();
                        }
                    });
                }

                //Update Page Dropdown button keydown event
                sizePerPageBtnElement.addEventListener("keydown", (evt: any) => {
                    if (evt.shiftKey && evt.key === stringsConstants.stringTab || evt.key === stringsConstants.stringEscape) {
                        const sizePerPageUlElement = this.getUlElement();
                        sizePerPageUlElement.setAttribute("style", "display:none");
                    }
                    else if (evt.keyCode === 40) {
                        const sizePerPageUlElement = this.getUlElement();
                        sizePerPageUlElement.setAttribute("style", "display:block");
                        pageSizeAnchorElements[0]?.focus();
                    }
                });

                //Update Page Size callout's first element keydown event
                const firstPageSizeElement = pageSizeAnchorElements[0];
                firstPageSizeElement.addEventListener("keydown", (evt: any) => {
                    if (evt.keyCode === 38) {
                        sizePerPageUlElement.setAttribute("style", "display:none");
                        sizePerPageBtnElement?.focus();
                    }
                    else if (evt.shiftKey && evt.key === stringsConstants.stringTab) {
                        sizePerPageUlElement.setAttribute("style", "display:none");
                    }
                    else if (evt.key === stringsConstants.stringTab || evt.shiftKey) {
                        sizePerPageUlElement.setAttribute("style", "display:block");
                    }
                });

                //Update Page Size callout's last element keydown event
                const lastPageSizeElement = pageSizeAnchorElements[pageSizeAnchorElements.length - 1];
                const paginationFirstBtn = this.totReportGridRef?.current?.querySelector(".pagination").getElementsByTagName('a')[0];
                lastPageSizeElement.addEventListener("keydown", (evt: any) => {
                    if (evt.keyCode === 40) {
                        sizePerPageUlElement.setAttribute("style", "display:none");
                        paginationFirstBtn.focus();
                    }
                    else if (!evt.shiftKey && evt.key === stringsConstants.stringTab) {
                        sizePerPageUlElement.setAttribute("style", "display:none");
                    }
                });
            }

            /**Update aria-expanded attribute in combobox and update focus for combobox list for Accessibility in Android **/
            if (prevState.tournamentsList !== this.state.tournamentsList && this.state.tournamentsList.length > 0) {
                //Update aria-expanded attribute in combobox for Accessibility in Android
                if (navigator.userAgent.match(/Android/i)) {

                    //Outside Click event for Report Tournaments combobox for Accessibility in Android
                    document.addEventListener("click", this.onReportTournamentsComboboxOutsideClick);

                    //remove aria-expanded attribute from combobox input element
                    const comboboxInput = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listbox-input");
                    comboboxInput.removeAttribute("aria-expanded");

                    //Update aria-expanded attribute for combobox expand/collapse button
                    const comboboxButton = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listboxwrapper").querySelector("button");
                    comboboxButton.setAttribute("aria-expanded", "false");

                    //get focus area combobox list wrapper element to set focus and attributes
                    const ulList: any = this.totReportComboboxWrapperRef?.current?.querySelector("#report-tournaments-listbox-list");
                    ulList.setAttribute("tabindex", "0");
                    comboboxButton.addEventListener("click", () => {
                        setTimeout(() => {
                            ulList.focus();
                        }, 1000);
                    });
                }
            }
        }
        catch (error) {
            console.error("CMP_TOT_TOTReport_componentDidUpdate \n", error);
        }
    }

    /**Remove Document click event listener on Unmount of Component **/
    public componentWillUnmount(): void {
        // For Accessibility in Android
        if (navigator.userAgent.match(/Android/i)) {
            document.removeEventListener("click", this.onReportTournamentsComboboxOutsideClick);
        }
        document.removeEventListener("click", this.onPageDropdownBtnOutsideClick);
    }

    //Close Report Tournaments Combobox Callout on click of outside for Accessibility in Android
    public onReportTournamentsComboboxOutsideClick = (evt: any) => {
        const isComboboxElement = document.getElementById("report-tournaments-listbox").contains(evt.target);
        if (!isComboboxElement) {
            this.mainComboboxRef.current.dismissMenu();
        }
    }

    //Get Page Dropdown Button's Callout Element from DOM
    public getUlElement = () => {
        const ulElements: any = this.totReportGridRef?.current?.getElementsByTagName("ul");
        let sizePerPageUlElement: HTMLUListElement;
        for (let ulElement of ulElements) {
            if (ulElement?.getAttribute("aria-labelledby") === "pageDropDown") {
                sizePerPageUlElement = ulElement;
                break;
            }
        }
        return sizePerPageUlElement;
    }

    //Close Size Per Page List on click of outside 
    public onPageDropdownBtnOutsideClick = (evt: any) => {
        const isBtnElement = evt?.target?.getAttribute("id") === "pageDropDown";
        if (!isBtnElement) {
            const sizePerPageUlElement = this.getUlElement();
            sizePerPageUlElement.setAttribute("style", "display:none");
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
                        totalCompletionPercentage: isNaN(totalPercentage) ? 0 : totalPercentage
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
    private async loadTopParticipants() {
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
                let organizedParticipants = commonServiceManager.groupBy(allParticipantsArray, (item: any) => item.User_x0020_Name);

                //Sum up the points for each participant
                organizedParticipants.forEach((participant) => {
                    let pointsCompleted: number = participant.reduce((previousValue: any, currentValue: any) => { return previousValue + currentValue["Points"]; }, 0);

                    //Push the metrics of each participant into an array.
                    topParticipantsArray.push({
                        Title: participant[0].User_x0020_Name,
                        Points: pointsCompleted
                    });
                });

                //Sort by points and then by user name 
                topParticipantsArray.sort((a: any, b: any) => {
                    if (a.Points < b.Points) return 1;
                    if (a.Points > b.Points) return -1;
                    if (a.Title > b.Title) return 1;
                    if (a.Title < b.Title) return -1;
                });

                //Get Top 5 participants and add it to Top Participants List
                let top5ParticipantsArray = topParticipantsArray.filter((item: any, idx: number) => idx < 5).map((item: any) => { return item; });

                if (top5ParticipantsArray.length > 0) {
                    //Loop through top tournaments and add the lables and data for the chart display
                    top5ParticipantsArray.forEach((element: any) => {
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
    private async loadTopTournaments() {
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
    private async loadParticipantsStatus() {
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
    private doughNutChartOptions: any = {
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

    //Set Pagination Properties
    private pagination = paginationFactory({
        page: 1,
        sizePerPage: 10,
        showTotal: true,
        alwaysShowAllBtns: false,
        //Render Page Size Options
        sizePerPageOptionRenderer: (options: any) => {
            return (
                <li className="dropdown-item" key={options.text} role="presentation" tabIndex={-1}>
                    <a
                        href="#"
                        role="menuitem"
                        tabIndex={0}
                        data-page={options.page}
                        onClick={() => {
                            options.onSizePerPageChange(options.page);
                            const sizePerPageUlElement = this.getUlElement();
                            sizePerPageUlElement.setAttribute("style", "display:none");
                            const sizePerPageBtnElement: HTMLButtonElement = this.totReportGridRef?.current?.querySelector("#pageDropDown");
                            sizePerPageBtnElement.setAttribute("aria-label", stringsConstants.sizePerPageLabel + " " + options.text);
                            sizePerPageBtnElement?.focus();
                        }}
                        onKeyDown={(evt: any) => {
                            const sizePerPageUlElement = this.getUlElement();
                            const sizePerPageBtnElement: HTMLButtonElement = this.totReportGridRef?.current?.querySelector("#pageDropDown");
                            if (evt.key === stringsConstants.stringSpace) {
                                options.onSizePerPageChange(options.page);
                                sizePerPageBtnElement.setAttribute("aria-label", stringsConstants.sizePerPageLabel + " " + options.text);
                                sizePerPageUlElement.setAttribute("style", "display:none");
                                sizePerPageBtnElement?.focus();
                            }
                            else if (evt.key === stringsConstants.stringEscape) {
                                sizePerPageUlElement.setAttribute("style", "display:none");
                                sizePerPageBtnElement?.focus();
                            }
                        }}
                        aria-label={stringsConstants.sizePerPageLabel + " " + options.text}
                    >{options.text}</a>
                </li>
            );
        },
        //customized the render options for page list in the pagination
        pageButtonRenderer: (options: any) => {
            const handleClick = (e: any) => {
                e.preventDefault();
                if (options.disabled) return;
                options.onPageChange(options.page);
            };
            const className = `${options.active ? 'active ' : ''}${options.disabled ? 'disabled ' : ''}`;
            let ariaLabel = "";
            let pageText = "";
            switch (options.title) {
                case "first page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<<';
                    break;
                case "previous page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<';
                    break;
                case "next page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>';
                    break;
                case "last page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>>';
                    break;
                default:
                    ariaLabel = `Go to page ${options.title}`;
                    pageText = options.title;
                    break;
            }
            return (
                <li key={options.title} className={`${className} page-item`} role="presentation" title={ariaLabel}>
                    <a className="page-link" href="#" onClick={handleClick} role="button" aria-label={ariaLabel}>
                        <span aria-hidden="true">{pageText}</span>
                    </a>
                </li>
            );
        },
        paginationTotalRenderer: (from: any, to: any, size: any) => {
            const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} results` : ""
            return (<span className="react-bootstrap-table-pagination-total" aria-live="polite" role="status">
                &nbsp;{resultsFound}
            </span>
            );
        }
    });

    /** On menu open add the attributes to fix the position issue in IOS and 
   Update aria-expanded attribute in combobox in Android for Accessibility **/
    private onMenuOpen = (listboxId: string) => {
        //Update aria-expanded attribute in combobox for Accessibility in Android
        if (navigator.userAgent.match(/Android/i)) {
            //remove aria-expanded attribute from combobox input element
            const comboboxInput = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listbox-input");
            comboboxInput.removeAttribute("aria-expanded");

            //Update aria-expanded attribute for combobox expand/collapse button
            const comboboxButton = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listboxwrapper").querySelector("button");
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

    // format the cell for participant Name
    participantFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        return (
            <Person
                personQuery={cell}
                view={3}
                personCardInteraction={1}
                className='participant-person-card'
            />
        );
    }

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
                formatter: this.participantFormatter
            }, {
                dataField: stringsConstants.ActivitiesCompleted,
                text: LocaleStrings.ActivitiesCompletedLabel,
                headerTitle: true,
                title: true,
                searchable: false
            }, {
                dataField: stringsConstants.Points,
                text: LocaleStrings.PointsLabel,
                sort: true,
                headerTitle: true,
                title: true,
                searchable: false
            }, {
                dataField: stringsConstants.TournamentCompletedPercentage,
                text: LocaleStrings.PercentageTournamentCompletedLabel,
                headerTitle: true,
                title: true,
                formatter: this.progressBar,
                searchable: false
            }];
        const { SearchBar } = Search;
        return (
            <>
                <div className={styles.container}>
                    <div className={styles.totReportPath}>
                        <img src={require("../assets/TOTImages/BackIcon.png")}
                            className={styles.backImg}
                            alt={LocaleStrings.BackButton}
                            aria-hidden="true"
                        />
                        <span
                            className={styles.backLabel}
                            onClick={() => this.props.onClickCancel()}
                            role="button"
                            tabIndex={0}
                            onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter || evt.key === stringsConstants.stringSpace) this.props.onClickCancel() }}
                            aria-label={LocaleStrings.TOTBreadcrumbLabel}
                        >
                            <span title={LocaleStrings.TOTBreadcrumbLabel}>
                                {LocaleStrings.TOTBreadcrumbLabel}
                            </span>
                        </span>
                        <span className={styles.border} aria-live="polite" role="alert" aria-label={LocaleStrings.TournamentReportsPageTitle + " Page"} />
                        <span className={styles.totReportLabel}>{LocaleStrings.TournamentReportsPageTitle}</span>
                    </div>
                    <br />
                    {this.state.noTournamentsFlag ?
                        <div>{LocaleStrings.NoCompletedTournamentsMessage}</div>
                        :
                        <div>
                            {this.state.tournamentsList.length > 0 && (
                                <div className={styles.totReportComboboxWrapper} ref={this.totReportComboboxWrapperRef}>
                                    <Label className={styles.totReportComboboxLabel}>{LocaleStrings.TournamentLabel} :</Label>
                                    <ComboBox
                                        defaultSelectedKey={stringsConstants.AllLabel}
                                        options={this.state.tournamentsList}
                                        onChange={this.setSelectedTournament.bind(this)}
                                        className={styles.totReportDropdown}
                                        calloutProps={{
                                            className: "totReportComboBoxCallout", directionalHintFixed: true, doNotLayer: true,
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
                                        id="report-tournaments-listbox"
                                        onMenuOpen={() => this.onMenuOpen("report-tournaments-listbox")}
                                        useComboBoxAsMenuWidth={true}
                                        ariaLabel={LocaleStrings.TournamentLabel}
                                        onMenuDismissed={() => {
                                            //Update aria-expanded attribute in combobox for Accessibility in Android
                                            if (navigator.userAgent.match(/Android/i)) {
                                                //remove aria-expanded attribute from combobox input element
                                                const comboboxInput = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listbox-input");
                                                comboboxInput.removeAttribute("aria-expanded");

                                                //Update aria-expanded attribute for combobox expand/collapse button
                                                const comboboxButton = this.totReportComboboxWrapperRef.current.querySelector("#report-tournaments-listboxwrapper").querySelector("button");
                                                comboboxButton.setAttribute("aria-expanded", "false");
                                            }
                                        }}
                                        componentRef={this.mainComboboxRef}
                                    />
                                    <div className={styles.dDInfoIconArea}>
                                        <TooltipHost
                                            content={LocaleStrings.ReportsDropdownInfoIconText}
                                            delay={window.innerWidth < stringsConstants.MobileWidth ? 0 : 2}
                                            directionalHint={DirectionalHint.topCenter}
                                        >
                                            <Icon iconName='Info' className={styles.dDInfoIcon} aria-label={LocaleStrings.ReportsDropdownInfoIconText} />
                                        </TooltipHost>
                                    </div>
                                </div>
                            )}
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
                                                className={styles.totReportChart}
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
                                                className={styles.totReportChart}
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
                                                className={styles.totReportChart}
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
                                                    <h2 className={styles.tableHeading} tabIndex={0} role="heading">{LocaleStrings.ParticipantsDetailsLabel}</h2>
                                                    <ExportCSVButton {...props.csvProps} className={styles.csvLink}>
                                                        <img
                                                            src={require("../assets/TOTImages/ExcelIcon.png")}
                                                            alt="Export Icon"
                                                            className={styles.excelIcon}
                                                        />
                                                        <span className={styles.downloadText}>{LocaleStrings.DownloadButtonText}</span>
                                                    </ExportCSVButton>
                                                </div>
                                                <div className={styles.participantsGrid} ref={this.totReportGridRef}>
                                                    <BootstrapTable
                                                        {...props.baseProps}
                                                        table-responsive={true}
                                                        defaultSorted={[{ dataField: stringsConstants.Points, order: 'desc' }]}
                                                        pagination={this.pagination}
                                                        noDataIndication={() => (<div aria-live="polite" role="alert">{LocaleStrings.NoRecordsinGridLabel}</div>)}
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


