import { Checkbox, Icon, Label, SearchBox, Spinner, SpinnerSize } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import React, { Component } from 'react';
import BootstrapTable from 'react-bootstrap-table-next';
import paginationFactory from 'react-bootstrap-table2-paginator';
import ToolkitProvider, { ToolkitContextType } from 'react-bootstrap-table2-toolkit';
import commonServices from '../Common/CommonServices';
import * as stringConstants from "../constants/strings";
import styles from "../scss/ManageApprovals.module.scss";

export interface IChampionPendingEvent {
    Champion: string;
    Email: string;
    Event: string;
    Date: Date;
    Points: number;
    Notes: string;
    EventActivityId: number;
}
export interface ChampionsActivitiesState {
    championPendingEvents: IChampionPendingEvent[];
    filteredChampionPendingEvents: IChampionPendingEvent[];
    showSpinner: boolean;
    selectedIds: Array<number>;
    isAllSelected: boolean;
    rejectMessage: string;
    approveMessage: string;
}

export interface ChampionsActivitiesProps {
    context: WebPartContext;
    siteUrl: string;
    setState: Function;
}
export default class ChampionsActivities extends Component<ChampionsActivitiesProps, ChampionsActivitiesState> {
    private commonServiceManager: commonServices;
    constructor(props: ChampionsActivitiesProps) {
        super(props);
        this.state = {
            championPendingEvents: [],
            filteredChampionPendingEvents: [],
            showSpinner: false,
            selectedIds: [],
            isAllSelected: false,
            rejectMessage: "",
            approveMessage: ""
        };
        //Create object for CommonServices class
        this.commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );

        this.selectChampionEvents = this.selectChampionEvents.bind(this);
    }

    //Life cycle method - componentDidMount
    public componentDidMount() {
        this.getChampionPendingEvents();
    }

    //Component did update life cycle method
    public componentDidUpdate(prevProps: Readonly<ChampionsActivitiesProps>, prevState: Readonly<ChampionsActivitiesState>, snapshot?: any): void {
        //updating state of the parent component 'ManageApprovals" to show/hide the notification icon
        if (prevState.championPendingEvents !== this.state.championPendingEvents) {
            if (this.state.championPendingEvents.length === 0) {
                this.props.setState({
                    isPendingEventApproval: false
                });
            }
            else {
                this.props.setState({
                    isPendingEventApproval: true
                });

            }
        }

    }

    //Get all champion pending events from event track details list.
    public async getChampionPendingEvents() {
        try {
            this.setState({ showSpinner: true });
            //Getting the pending events from event track details list
            let pendingEventsFilterQuery = "Status eq '" + stringConstants.pendingStatus + "'";
            const sortColumn = "Created";
            const pendingEvents: any[] = await this.commonServiceManager.getItemsSortedWithFilter(stringConstants.EventTrackDetailsList, pendingEventsFilterQuery, sortColumn);

            //Getting member ids from event track details list
            const memberIds = [];
            pendingEvents.forEach((event: any) => memberIds.push(event.MemberId));

            //Filtering member ids from duplicate member ids
            const filteredMemberIds = memberIds.filter((item, index) => memberIds.indexOf(item) === index);

            //Creating a filter query to fetch data from member list
            let memberIdFilterQuery = "";
            filteredMemberIds.forEach((id, idx) => {
                memberIdFilterQuery = idx === 0 ? "ID eq " + id : memberIdFilterQuery + " or ID eq " + id;
            });
            //Getting data from member list
            const memberData: any[] = await this.commonServiceManager.getItemsWithOnlyFilter(stringConstants.MemberList, memberIdFilterQuery);

            //Mapping events and member email id
            let championEvents: IChampionPendingEvent[] = [];
            pendingEvents.forEach((event) => {
                const activity: IChampionPendingEvent = {
                    EventActivityId: event.ID,
                    Champion: event.MemberName ? event.MemberName : "",
                    Email: "",
                    Event: event.EventName ? event.EventName : "",
                    Date: event.DateofEvent ? new Date(event.DateofEvent) : new Date(),
                    Points: event.Count ? event.Count : "",
                    Notes: event.Notes ? event.Notes : ""
                };
                memberData.forEach((data) => {
                    if (data.ID === event.MemberId) {
                        activity.Email = data.Title ? data.Title : "";
                    }
                });
                championEvents.push(activity);
            });

            this.setState({ championPendingEvents: championEvents, showSpinner: false });
        }
        catch (error) {
            console.error("CMP_ChampionActivities_getChampionPendingEvents \n", JSON.stringify(error));
        }
    }

    //Method to update the status in the Event Track details List
    private async updatePendingEvents(statusText: string, selectedIDs: any) {
        try {
            this.setState({
                showSpinner: true,
                rejectMessage: "",
                approveMessage: ""
            });

            let updateEventObject: any = {
                Status: statusText
            };

            //Update status for pending items in Member List
            let updateResponse = await this.commonServiceManager.updateMultipleItems(stringConstants.EventTrackDetailsList, updateEventObject, selectedIDs);

            //Refresh the items shown in the grid
            await this.getChampionPendingEvents();

            if (updateResponse) {
                //Updating state variables based on the approval action
                if (statusText === stringConstants.approvedStatus) {
                    this.setState({
                        approveMessage: LocaleStrings.EventApprovedMessage,
                        selectedIds: [],
                        rejectMessage: "",
                        isAllSelected: false,
                        showSpinner: false
                    });
                } else if (statusText === stringConstants.rejectedStatus) {
                    this.setState({
                        rejectMessage: LocaleStrings.EventRejectedMessage,
                        selectedIds: [],
                        approveMessage: "",
                        isAllSelected: false,
                        showSpinner: false
                    });
                }
            } else {
                //If any error occurs during batch update
                this.setState({
                    rejectMessage: stringConstants.CMPErrorMessage + "while approving/rejecting champion event request(s).",
                    selectedIds: [],
                    isAllSelected: false,
                    showSpinner: false
                });
            }
        }
        catch (error) {
            //Refresh the items shown in the grid
            await this.getChampionPendingEvents();

            this.setState({
                rejectMessage: stringConstants.CMPErrorMessage + "while approving/rejecting champion event request(s). Below are the details: \n" + JSON.stringify(error),
                selectedIds: [],
                isAllSelected: false,
                showSpinner: false
            });
            console.error("CMP_ChampionActivities_updatePendingEvents \n", JSON.stringify(error));
        }
    }

    //Update all selected events to new array
    public selectChampionEvents(isChecked: boolean, key: number, selectAll: boolean) {
        //When "Select All" is checked
        if (selectAll && isChecked) {
            this.setState({ isAllSelected: true });
            let selectedEvents = [];
            this.state.filteredChampionPendingEvents.forEach((event: IChampionPendingEvent) => {
                selectedEvents.push(event.EventActivityId);
            });
            this.setState({ selectedIds: selectedEvents });
        }
        // When "Select All" is unchecked
        else if (selectAll && !isChecked) {
            this.setState({ isAllSelected: false, selectedIds: [] });
        }
        else {
            //When checkbox is checked
            if (isChecked) {
                let selectedEvents = this.state.selectedIds;
                selectedEvents.push(key);
                this.setState({ selectedIds: selectedEvents });

                //Automatically check the "Select All" option when the last checkbox is checked
                if (selectedEvents.length === this.state.filteredChampionPendingEvents.length) {
                    this.setState({ isAllSelected: true });
                }

            }
            //When checkbox is unchecked
            else {
                const selectedEvents = this.state.selectedIds.filter((eventId: any) => {
                    return eventId !== key;
                });
                this.setState({
                    isAllSelected: false,
                    selectedIds: selectedEvents
                });
            }
        }
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

    public render() {

        const eventsTableHeader = [
            {
                dataField: "EventActivityId",
                headerTitle: () => LocaleStrings.SelectAllEvents,
                headerFormatter: () => {
                    return (
                        <Checkbox
                            onChange={(_eve: any, isChecked: boolean) => {
                                this.selectChampionEvents(isChecked, -1, true);
                            }}
                            className={styles.selectAllCheckbox}
                            checked={this.state.isAllSelected}
                            ariaLabel={LocaleStrings.SelectAllEvents}
                            disabled={this.state.showSpinner || this.state.championPendingEvents.length === 0}
                        />
                    );
                },
                title: () => LocaleStrings.SelectEvent,
                attrs: (_cell: any, row: any) => ({ key: row.EventActivityId }),
                formatter: (_: any, gridRow: any) => {
                    return (
                        <Checkbox
                            onChange={(_eve: any, isChecked: boolean) => {
                                this.selectChampionEvents(isChecked, gridRow.EventActivityId, false);
                            }}
                            className={styles.selectItemCheckbox}
                            ariaLabel={LocaleStrings.SelectEvent}
                            checked={this.state.selectedIds.includes(gridRow.EventActivityId)}
                            disabled={this.state.showSpinner}
                        />
                    );
                },
                searchable: false
            },
            {
                dataField: "Champion",
                text: LocaleStrings.ChampionLabel,
                headerTitle: LocaleStrings.ChampionLabel,
                title: true,
                searchable: true,
                sort: true
            },
            {
                dataField: "Email",
                text: LocaleStrings.EmailLabel,
                headerTitle: LocaleStrings.EmailLabel,
                title: true,
                searchable: true,
                sort: true
            },
            {
                dataField: "Event",
                text: LocaleStrings.EventLabel,
                headerTitle: LocaleStrings.EventLabel,
                title: true,
                searchable: true
            },
            {
                dataField: "Date",
                text: LocaleStrings.DateLabel,
                headerTitle: LocaleStrings.DateLabel,
                title: (_cell: any, row: any) => row.Date.toDateString().slice(4),
                searchable: false,
                sort: true,
                formatter: (_cell: any, gridRow: any) => <>{gridRow.Date.toDateString().slice(4)}</>
            },
            {
                dataField: "Points",
                text: LocaleStrings.PointsLabel,
                headerTitle: LocaleStrings.PointsLabel,
                title: true,
                searchable: false
            },
            {
                dataField: "Notes",
                text: LocaleStrings.NotesLabel,
                headerTitle: LocaleStrings.NotesLabel,
                title: true,
                searchable: false
            }
        ];

        return (
            <div className={styles.approvalsContainer}>
                {this.state.approveMessage &&
                    <Label className={styles.approveMessage + ' col-xl-5 col-lg-5 col-md-6 col-sm-8 col-xs-9'}>
                        <img
                            src={require('../assets/TOTImages/tickIcon.png')}
                            alt={LocaleStrings.SuccessIcon}
                            className={styles.tickImage}
                        />
                        {this.state.approveMessage}
                    </Label>
                }
                {this.state.rejectMessage &&
                    <Label className={styles.rejectMessage + ' col-xl-5 col-lg-5 col-md-6 col-sm-8 col-xs-9'}>
                        {this.state.rejectMessage}
                    </Label>
                }
                <ToolkitProvider
                    bootstrap4
                    keyField="EventActivityId"
                    data={this.state.championPendingEvents}
                    columns={eventsTableHeader}
                    search={{
                        afterSearch: (newResult: IChampionPendingEvent[]) => {
                            this.setState({
                                filteredChampionPendingEvents: newResult,
                                isAllSelected: newResult.length === this.state.selectedIds.length ? true : false
                            });
                        }
                    }}
                >
                    {
                        (props: ToolkitContextType) => (
                            <div>
                                {this.state.championPendingEvents.length > 0 &&
                                    <>
                                        <div className={'col-xl-5 col-lg-5 col-md-6 col-sm-7 col-xs-9' + " " + styles.searchboxPadding}>
                                            <SearchBox
                                                placeholder={LocaleStrings.PendingEventsSearchboxPlaceholder}
                                                onChange={(_, searchedText) => props.searchProps.onSearch(searchedText)}
                                                className={styles.approvalsSearchbox}
                                            />
                                        </div>
                                        {this.state.selectedIds.length > 0 &&
                                            <Label className={styles.selectedRowText}>
                                                {this.state.selectedIds.length} {LocaleStrings.EventsSelectedLabel}
                                            </Label>
                                        }
                                    </>
                                }
                                <div>
                                    <BootstrapTable
                                        striped
                                        {...props.baseProps}
                                        table-responsive={true}
                                        pagination={(this.state.filteredChampionPendingEvents.length > 0 && this.state.filteredChampionPendingEvents.length > 10) && this.pagination}
                                        wrapperClasses={styles.approvalsTableWrapper}
                                        headerClasses={styles.eventsApprovalTableHeader}
                                        bodyClasses={styles.eventsApprovalTableBody}
                                        noDataIndication={() => (
                                            <div className={styles.noRecordsArea}>
                                                {this.state.showSpinner ?
                                                    <Spinner
                                                        label={LocaleStrings.ProcessingSpinnerLabel}
                                                        size={SpinnerSize.large}
                                                    /> :
                                                    <>
                                                        <img
                                                            src={require('../assets/CMPImages/Norecordsicon.svg')}
                                                            alt={LocaleStrings.NoRecordsIcon}
                                                            className={styles.noRecordsImg}
                                                        />
                                                        <span className={styles.noRecordsLabels}>
                                                            {this.state.championPendingEvents.length === 0 ?
                                                                LocaleStrings.NoPendingEventsLabel
                                                                :
                                                                LocaleStrings.NoSearchResults
                                                            }
                                                        </span>
                                                    </>
                                                }
                                            </div>
                                        )}
                                    />
                                </div>
                            </div>
                        )
                    }
                </ToolkitProvider>
                {this.state.showSpinner && this.state.championPendingEvents.length > 0 &&
                    <Spinner
                        label={LocaleStrings.ProcessingSpinnerLabel}
                        size={SpinnerSize.large}
                    />
                }
                {this.state.championPendingEvents.length > 0 &&
                    <div className={styles.manageApprovalsBtnArea}>
                        <button
                            className={`btn ${styles.approveBtn}`}
                            onClick={e => this.updatePendingEvents(stringConstants.approvedStatus, this.state.selectedIds)}
                            title={LocaleStrings.ApproveButton}
                            disabled={this.state.selectedIds.length === 0}
                        >
                            <Icon iconName="Completed" className={styles.approveBtnIcon} />
                            <span className={styles.approveBtnLabel}>{LocaleStrings.ApproveButton}</span>
                        </button>
                        <button
                            className={"btn " + styles.rejectBtn}
                            onClick={e => this.updatePendingEvents(stringConstants.rejectedStatus, this.state.selectedIds)}
                            title={LocaleStrings.RejectButton}
                            disabled={this.state.selectedIds.length === 0}
                        >
                            <Icon iconName="ErrorBadge" className={styles.rejectBtnIcon} />
                            <span className={styles.rejectBtnLabel}>{LocaleStrings.RejectButton}</span>
                        </button>
                    </div>
                }
            </div>

        );
    }
}
