import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
//React Boot Strap
import BootstrapTable from "react-bootstrap-table-next";
import paginationFactory from "react-bootstrap-table2-paginator";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
//FluentUI controls
import { IButtonStyles, DefaultButton } from "@fluentui/react";
import { Icon, IIconProps } from '@fluentui/react/lib/Icon';
import { Label } from "@fluentui/react/lib/Label";
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTLeaderBoard.module.scss";
import TOTSidebar from "./TOTSideBar";
import { RxJsEventEmitter } from "../events/RxJsEventEmitter";
import { EventData } from "../events/EventData";
import * as LocaleStrings from 'ClbHomeWebPartStrings';

//Global variables
let commonService: commonServices;
let currentUserEmail: string = "";

const columns = [
  {
    dataField: "Rank",
    text: LocaleStrings.RankLabel,
  },
  {
    dataField: "User",
    text: LocaleStrings.UserLabel,
  },
  {
    dataField: "Points",
    text: LocaleStrings.PointsLabel,
  },
];

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', cursor: 'pointer' } };

const calloutProps = { gapSpace: 0 };

const classes = mergeStyleSets({
  icon: {
    fontSize: '16px',
    paddingLeft: '10px',
    fontWeight: 'bolder',
    color: '#1d0f62',
    position: 'relative',
    top: '3px'
  }
});
const backIcon: IIconProps = { iconName: 'NavigateBack' };

const backBtnStyles: Partial<IButtonStyles> = {
  root: {
    borderColor: "#33344A",
    backgroundColor: "white",
    marginLeft: "1.5%"
  },
  rootHovered: {
    borderColor: "#33344A",
    backgroundColor: "white",
    color: "#000003"
  },
  rootPressed: {
    borderColor: "#33344A",
    backgroundColor: "white",
    color: "#000003"
  },
  icon: {
    fontSize: "17px",
    fontWeight: "bolder",
    color: "#000003",
    opacity: 1
  },
  label: {
    font: "normal normal bold 14px/24px Segoe UI",
    letterSpacing: "0px",
    color: "#000003",
    opacity: 1,
    marginTop: "-3px"
  }
};

export interface ITOTLeaderBoardProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
  onClickMyDashboardLink: Function;
}
interface ITOTLeaderBoardState {
  showSuccess: Boolean;
  showError: Boolean;
  noActiveParticipants: boolean;
  noActiveTournament: boolean;
  errorMessage: string;
  tournamentName: any;
  tournamentDescription: string;
  activeTournamentsList: Array<any>;
  myTournamentsList: Array<any>;
  activeTournamentName: any;
  myTournamentName: any;
  tournamentDescriptionList: Array<any>;
  allUserActions: any;
  isShowLoader: boolean;
  currentUserDetails: any;
  userLoaded: string;
}

export default class TOTLeaderBoard extends React.Component<
  ITOTLeaderBoardProps,
  ITOTLeaderBoardState
> {
  private readonly _eventEmitter: RxJsEventEmitter =
    RxJsEventEmitter.getInstance();
  constructor(props: ITOTLeaderBoardProps, state: ITOTLeaderBoardState) {
    super(props);
    //Set default values
    this.state = {
      showSuccess: false,
      showError: false,
      noActiveParticipants: false,
      noActiveTournament: false,
      errorMessage: "",
      tournamentName: "",
      tournamentDescription: "",
      activeTournamentsList: [],
      myTournamentsList: [],
      activeTournamentName: "",
      myTournamentName: "",
      tournamentDescriptionList: [],
      allUserActions: [],
      isShowLoader: false,
      currentUserDetails: [],
      userLoaded: "",
    };
    //Create object for commonServices class
    commonService = new commonServices(this.props.context, this.props.siteUrl);
    // Bind methods 
    this.getActiveTournamentActions = this.getActiveTournamentActions.bind(this);
    this.getMyTournamentActions = this.getMyTournamentActions.bind(this);
  }

  //Get User Actions from list and bind to table
  public componentDidMount() {
    this.setState({
      isShowLoader: true,
    });
    //Get list of active tournaments from Tournaments list
    this.getActiveTournaments();
  }

  //Get list of active tournaments from Tournaments list and binding it to dropdowns
  private async getActiveTournaments() {
    console.log(stringsConstants.TotLog + "Getting list of active tournaments from Tournaments list.");
    try {
      //Get current users's email
      currentUserEmail =
        this.props.context.pageContext.user.email.toLowerCase();

      //Get current active tournament details
      let activeTournamentDetails: any[] =
        await commonService.getActiveTournamentDetails();

      var activeTournamentsChoices = [];
      var myTournamentsChoices = [];
      var tournamentDescriptionChoices = [];
      //If active tournament found
      if (activeTournamentDetails.length > 0) {

        //Get current user's active tournament details
        let filterUserTournaments: string = "Title eq '" + currentUserEmail + "'";

        const currentUserTournaments: any[] =
          await commonService.getFilteredListItemsWithSpecificColumns(
            stringsConstants.UserActionsList, "Tournament_x0020_Name",
            filterUserTournaments
          );

        const uniqueUserTournaments: any[] = currentUserTournaments.filter((value, index) => {
          const _value = JSON.stringify(value);
          return index === currentUserTournaments.findIndex(item => {
            return JSON.stringify(item) === _value;
          });
        });

        //Loop through all "Active" tournaments and create an array with key and text
        await activeTournamentDetails.forEach((eachTournament) => {

          //Create an array for My Tournaments dropdown
          if (uniqueUserTournaments.some(tournament => tournament.Tournament_x0020_Name == eachTournament["Title"])) {
            myTournamentsChoices.push({
              key: eachTournament["Title"],
              text: eachTournament["Title"]
            });
          }
          else {
            //Create an array for Active Tournaments dropdown
            activeTournamentsChoices.push({
              key: eachTournament["Title"],
              text: eachTournament["Title"]
            });
          }
          tournamentDescriptionChoices.push({
            key: eachTournament["Title"],
            text: eachTournament["Description"]
          });
        });
        activeTournamentsChoices.sort((a, b) => a.text.localeCompare(b.text));
        myTournamentsChoices.sort((a, b) => a.text.localeCompare(b.text));

        //Set state variables for dropdown options
        this.setState({
          activeTournamentsList: activeTournamentsChoices,
          myTournamentsList: myTournamentsChoices,
          tournamentDescriptionList: tournamentDescriptionChoices,
          isShowLoader: false
        });

        //Set the first option as a default tournament for My Tournaments dropdown
        if (myTournamentsChoices.length > 0) {
          this.setState({
            myTournamentName: myTournamentsChoices[0].text,
            activeTournamentName: null,
            tournamentName: myTournamentsChoices[0].text,
          });
        }
        //If My Tournaments is empty, Set the first option as a default tournament for Active Tournaments dropdown
        else if (activeTournamentsChoices.length > 0) {
          this.setState({
            myTournamentName: null,
            activeTournamentName: activeTournamentsChoices[0].text,
            tournamentName: activeTournamentsChoices[0].text,
          });
        }
      }
      //If there is no active tournament
      else {
        this.setState({
          showError: true,
          errorMessage: LocaleStrings.NoActiveTournamentMessage,
          noActiveTournament: true,
          userLoaded: "1",
          isShowLoader: false
        });
      }
    }
    catch (error) {
      console.error("TOT_TOTMyDashboard_getActiveTournaments \n", error);
    }
  }

  //Set a value when an option is selected in My Tournaments dropdown and reset the Active Tournaments dropdown
  public getMyTournamentActions = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({
      tournamentName: option.key,
      myTournamentName: option.key,
      activeTournamentName: null
    });
  }

  //Set a value when an option is selected in Active Tournaments dropdown and reset the My Tournaments dropdown
  public getActiveTournamentActions = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({
      tournamentName: option.key,
      activeTournamentName: option.key,
      myTournamentName: null
    });
  }
  //Refresh the user details table whenever the tournament name is selected
  public componentDidUpdate(prevProps: Readonly<ITOTLeaderBoardProps>, prevState: Readonly<ITOTLeaderBoardState>, snapshot?: any): void {
    if (prevState.tournamentName != this.state.tournamentName) {
      if (this.state.tournamentName !== "")
        this.getUserActions();
      //Refresh the points and rank in the sidebar when a tournament is selected in My Tournaments / Active tournaments dropdown
      this._eventEmitter.emit("rebindSideBar:start", {
        tournamentName: this.state.tournamentName,
      } as EventData);
    }
  }

  //Get all user action for active tournament and bind to table
  private async getUserActions(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {

        //Set the description for selected tournament
        var tournmentDesc = this.state.tournamentDescriptionList.find((item) => item.key == this.state.tournamentName);

        this.setState({
          tournamentDescription: tournmentDesc.text
        });
        //get active tournament's participants
        await commonService
          .getUserActions(this.state.tournamentName)
          .then((res) => {
            if (res.length > 0) {
              this.setState({
                allUserActions: res,
                isShowLoader: false,
                userLoaded: "1",
              });
            } else if (res.length == 0) {
              this.setState({
                allUserActions: [],
                userLoaded: "0",
                showError: true,
                noActiveParticipants: true,
                errorMessage: LocaleStrings.NoActiveParticipantsMessage,
                isShowLoader: false,
              });
            } else if (res == "Failed") {
              console.error("TOT_TOTLeaderboard_getUserActions \n");
            }
          });

      } catch (error) {
        console.error("TOT_TOTLeaderboard_getUserActions \n", error);
        this.setState({
          showError: true,
          errorMessage:
            stringsConstants.TOTErrorMessage +
            " while getting user actions. Below are the details: \n" +
            JSON.stringify(error),
          showSuccess: false,
          isShowLoader: false,
        });
      }
    });
  }

  public render(): React.ReactElement<ITOTLeaderBoardProps> {
    return (
      <div>
        {this.state.isShowLoader && <div className={styles.load}></div>}
        <div className={styles.container}>
          <div className={styles.totSideBar}>
            {this.state.userLoaded != "" && (
              <TOTSidebar
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                currentUserDetails={this.state.allUserActions}
                onClickCancel={() => this.props.onClickCancel()}
              />
            )}
            <div className={styles.contentTab}>
              <div className={styles.totLeaderboardPath}>
                <img src={require("../assets/CMPImages/BackIcon.png")}
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
                <span className={styles.totLeaderboardLabel}>{LocaleStrings.TOTLeaderBoardPageTitle}</span>
              </div>
              <div className={styles.dropdownArea}>
                <Row>
                  {this.state.myTournamentsList.length > 0 && (
                    <Col md={5}>
                      <span className={styles.labelHeading}>{LocaleStrings.MyTournamentsLabel} :
                        <TooltipHost
                          content={LocaleStrings.MyTournamentsTooltip}
                          calloutProps={calloutProps}
                          styles={hostStyles}
                        >
                          <Icon aria-label="Info" iconName="Info" className={classes.icon} />
                        </TooltipHost>
                      </span>
                      <ComboBox className={styles.dropdownCol}
                        placeholder={LocaleStrings.SelectTournamentPlaceHolder}
                        selectedKey={this.state.myTournamentName}
                        options={this.state.myTournamentsList}
                        onChange={this.getMyTournamentActions.bind(this)}
                        ariaLabel={LocaleStrings.MyTournamentsLabel + ' list'}
                      />
                    </Col>
                  )}
                  {this.state.myTournamentsList.length > 0 && this.state.activeTournamentsList.length > 0 && (
                    <Col md={1} className={styles.labelCol} >
                      <span className={styles.labelHeading}>{LocaleStrings.OrLabel}</span>
                    </Col>
                  )}
                  {this.state.activeTournamentsList.length > 0 && (
                    <Col md={5}>
                      <span className={styles.labelHeading}>{LocaleStrings.ActiveTournamentLabel} : </span>
                      <ComboBox className={styles.dropdownCol}
                        placeholder={LocaleStrings.SelectTournamentPlaceHolder}
                        selectedKey={this.state.activeTournamentName}
                        options={this.state.activeTournamentsList}
                        onChange={this.getActiveTournamentActions.bind(this)}
                        ariaLabel={LocaleStrings.ActiveTournamentLabel + " list"}
                      />
                    </Col>
                  )}
                </Row>
              </div>
              {this.state.tournamentName != "" && (
                <div>
                  {this.state.tournamentName != "" && (
                    <ul className={styles.listArea}>
                      {this.state.tournamentDescription && (
                        <li className={styles.listVal}>
                          <span className={styles.labelHeading}>
                            {LocaleStrings.DescriptionLabel}
                          </span>
                          :
                          <span className={styles.labelNormal}>
                            {this.state.tournamentDescription}
                          </span>
                        </li>
                      )}
                    </ul>
                  )}
                  <div className={styles.table}>
                    {this.state.allUserActions.length > 0 ? (
                      <BootstrapTable
                        table-responsive
                        bordered
                        hover
                        keyField="Rank"
                        data={this.state.allUserActions}
                        columns={columns}
                        pagination={paginationFactory()}
                        headerClasses="header-class"
                      />
                    )
                      :
                      <div>
                        {this.state.showError && this.state.noActiveParticipants && (
                          <Label className={styles.noActvPartErr}>
                            {LocaleStrings.NoActiveParticipantsErrorMessage}
                            <span className={styles.myDashboardLink}
                              onClick={() => this.props.onClickMyDashboardLink()}>
                              {LocaleStrings.TOTMyDashboardPageTitle}
                            </span>!
                          </Label>
                        )}
                      </div>
                    }
                  </div>
                </div>
              )}
              <div className={styles.contentArea}>
                {this.state.showError && !this.state.noActiveParticipants && (
                  <Label className={this.state.noActiveTournament ? styles.noActvTourErr : styles.errorMessage}>
                    {this.state.errorMessage}
                  </Label>
                )}
              </div>
              <div>
                <DefaultButton
                  text={LocaleStrings.BackButton}
                  title={LocaleStrings.BackButton}
                  iconProps={backIcon}
                  onClick={() => this.props.onClickCancel()}
                  styles={backBtnStyles}>
                </DefaultButton>
              </div>
            </div>
          </div>
        </div>
      </div> //outer div
    ); //close return
  } //end of render
}
