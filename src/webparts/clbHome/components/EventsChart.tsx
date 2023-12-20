import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as React from 'react';
import * as stringsConstants from '../constants/strings';
import commonServices from '../Common/CommonServices';
import styles from '../scss/ChampionReport.module.scss';
import { Chart } from 'chart.js';
import { ChartControl, ChartType } from '@pnp/spfx-controls-react/lib/ChartControl';
import { Component } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '../scss/Champions.scss';

// Common Services Object
let commonServiceManager: commonServices;

// Vertical Bar Chart Variables and Values
const verticalChartColor: any = ['#02A5F2', '#FFBA02', '#888686', '#5BBB02', '#FF4F17', '#02A5F2', '#FFBA02', '#888686', '#5BBB02', '#FF4F17'];
let chartLabelLength: number;
let chartFontSize: number;
let chartStepValue: number;

export interface EventsChartProps {
    context: WebPartContext;
    siteUrl: string;
    callBack?: Function;
    filteredAllEvents: Array<any>;
    parentComponent: string;
    selectedMemberID: string;
    updateEventsListHeight?: Function;
    currentThemeName?: string;
}

export interface EventsChartState {
    topEventTypesChartdata: Chart.ChartData;
    isDesktop: boolean;
}

export default class EventsChart extends Component<EventsChartProps, EventsChartState> {
    private barChartRef: React.RefObject<HTMLDivElement>;
    constructor(props: any) {
        super(props);
        this.barChartRef = React.createRef();

        // States
        this.state = {
            topEventTypesChartdata: {},
            isDesktop: true
        };

        // Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );

    }
    // Method to load top 10 events for chart control
    public componentDidMount() {
        // Adding window resize event listener while mounting the component
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
        // Assign appropriate values to the chart variables
        chartLabelLength = this.state.isDesktop ? 14 : 8;
        chartFontSize = this.state.isDesktop ? 14 : 12;
        chartStepValue = this.state.isDesktop ? 5 : 10;

        this.loadTopEvents(this.props.selectedMemberID);
    }

    // Set the state object for screen size
    resize = () => {
        this.setState({
            isDesktop: this.props.parentComponent === stringsConstants.SidebarLabel ? ((window.innerWidth > 745 && window.innerWidth < 992)
                || (window.innerWidth > 1092)) :
                this.props.parentComponent === stringsConstants.ChampionReportLabel ? ((window.innerWidth > 540 && window.innerWidth <= 767)
                    || (window.innerWidth > 991)) : true
        });
        this.props.parentComponent === stringsConstants.ChampionReportLabel && this.props.updateEventsListHeight(this.barChartRef?.current.getElementsByTagName("div")[1].clientHeight + "px");
    };

    // Before unmounting, remove event listener
    componentWillUnmount() {
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // This method will be called whenever there is an update to the component
    public componentDidUpdate(prevProps: Readonly<EventsChartProps>, prevState: Readonly<EventsChartState>, snapshot?: any) {

        if (prevProps.selectedMemberID !== this.props.selectedMemberID ||
            prevProps.filteredAllEvents !== this.props.filteredAllEvents) {
            this.loadTopEvents(this.props.selectedMemberID);
        }
        if (prevState.isDesktop !== this.state.isDesktop) {
            chartLabelLength = this.state.isDesktop ? 14 : 8;
            chartFontSize = this.state.isDesktop ? 14 : 12;
            chartStepValue = this.state.isDesktop ? 5 : 10;
        }
    }


    // Get data for top 10 events
    private async loadTopEvents(selectedChampion: string) {
        try {
            let arrLabels: string[] = [];
            let arrData: number[] = [];
            let topEventsArray: any[] = [];
            let topEvents: any[] = [];
            this.setState({
                topEventTypesChartdata: {}
            });
            if (selectedChampion === stringsConstants.AllLabel) {
                if (this.props.filteredAllEvents.length > 0) {
                    let organizedEvents = commonServiceManager.groupBy(this.props.filteredAllEvents, (item: any) => item.EventName);

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

                    //get Top 10 Events
                    topEvents = topEventsArray.filter((item, idx) => idx < stringsConstants.ChartEventsLimit).map(item => { return item; });

                    if (topEvents.length > 0) {
                        //Loop through top events and add the labels and data for the chart display
                        topEvents.forEach(element => {
                            arrLabels.push(element.Title);
                            arrData.push(element.Count);
                        });
                    }
                }
            } else {
                if (this.props.filteredAllEvents.length > 0) {

                    let championEvents = this.props.filteredAllEvents.filter((item) => item.MemberId === selectedChampion);
                    if (championEvents.length > 0) {
                        let organizedEvents = commonServiceManager.groupBy(championEvents, (item: any) => item.EventName);

                        //count the number of events for each event type
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

                        //get Top 10 Events
                        topEvents = topEventsArray.filter((item, idx) => idx < stringsConstants.ChartEventsLimit).map(item => { return item; });

                        if (topEvents.length > 0) {
                            //Loop through top events and add the labels and data for the chart display
                            topEvents.forEach(element => {
                                arrLabels.push(element.Title);
                                arrData.push(element.Count);
                            });
                        }
                    }
                }
            }
            this.setState({
                topEventTypesChartdata: {
                    labels: arrLabels,
                    datasets: [{
                        label: '',
                        data: arrData,
                        backgroundColor: verticalChartColor,
                        fill: false
                    }]
                },
            });
            //Resetting the state to render the chart data with current state value
            setTimeout(() => {
                this.setState({
                    topEventTypesChartdata: this.state.topEventTypesChartdata,
                });
            }, 10);
        }
        catch (error) {
            console.error("CMP_EventsChart_loadTopEvents \n", error);
        }
    }

    // Main render method
    public render() {
        // To determine whether the component is called from sidebar or not
        const isSidebar = this.props.parentComponent === stringsConstants.SidebarLabel;
        const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
        // Options for vertical bar chart
        const verticalChartOptions = {
            legend: {
                labels: {
                    boxWidth: 0,
                }
            },
            scales: {
                yAxes: [{
                    stacked: true,
                    ticks: {
                        fontColor: isDarkOrContrastTheme ? "#FFFFFF" : "#020000",
                        fontSize: chartFontSize,
                        tooltips: true,
                        stepSize: chartStepValue,
                        callback: (label: string) => {
                            if (label.length > chartLabelLength) {
                                return label.slice(0, chartLabelLength) + '...';
                            } else {
                                return label;
                            }
                        }
                    }
                }],
                xAxes: [{
                    stacked: true,
                    ticks: {
                        fontColor: isDarkOrContrastTheme ? "#FFFFFF" : "#020000",
                        fontSize: chartFontSize,
                        stepSize: chartStepValue,
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
        return (
            <React.Fragment>
                <div className={`${styles.chartArea}${isSidebar ? " dashboard-chart-wrapper" : ""}`} ref={this.barChartRef}>
                    <div className={`${isSidebar ? " dashboard-chart-heading" : styles.chartHeading}`}>{LocaleStrings.EventChartLabel}</div>
                    <ChartControl
                        type={ChartType.Bar}
                        data={this.state.topEventTypesChartdata}
                        options={verticalChartOptions}
                        accessibility={{ enable: true, alternateText: `${LocaleStrings.EventChartLabel} chart` }}
                        className={`${styles.cmpReportChart}${isSidebar ? " sidebar-chart" : ""}`}
                    />
                </div>
            </React.Fragment>
        );
    }
}
