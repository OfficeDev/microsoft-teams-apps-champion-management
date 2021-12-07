//FluentUI controls
import { IButtonStyles, DefaultButton } from "@fluentui/react";
import { Icon, IIconProps } from '@fluentui/react/lib/Icon';
import { Label } from "@fluentui/react/lib/Label";
import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as React from "react";
//React Boot Strap
import BootstrapTable from "react-bootstrap-table-next";
import paginationFactory from "react-bootstrap-table2-paginator";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTLeaderBoard.module.scss";
import TOTSidebar from "./TOTSideBar";

//global variables
let commonService: commonServices;
let allUsersDetails: any = [];
const columns = [
  {
    dataField: "Rank",
    text: "Rank",
  },
  {
    dataField: "User",
    text: "User",
  },
  {
    dataField: "Points",
    text: "Points",
  },
];

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
  tournamentName: string;
  tournamentDescription: string;
  allUserActions: any;
  isShowLoader: boolean;
  currentUserDetails: any;
  userLoaded: string;
}

export default class TOTLeaderBoard extends React.Component<
  ITOTLeaderBoardProps,
  ITOTLeaderBoardState
> {
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
      allUserActions: [],
      isShowLoader: false,
      currentUserDetails: [],
      userLoaded: "",
    };
    //Create object for commonServices class
    commonService = new commonServices(this.props.context, this.props.siteUrl);
  }
  public _graphClient: MSGraphClient;

  //Get User Actions from list and bind to table
  public componentDidMount() {
    this.setState({
      isShowLoader: true,
    });
    this.getAllUsers().then((res) => {
      if (res == "Success") {
        this.getUserActions();
      }
    });
  }

  //get all users properties and store in array
  private async getAllUsers(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      this._graphClient =
        await this.props.context.msGraphClientFactory.getClient();
      await this._graphClient
        .api("/users")
        .get()
        .then(async (users: any, rawResponse?: any) => {
          for (let user of users.value) {
            if (user.mail != null) {
              allUsersDetails.push({
                email: user.mail.toLowerCase(),
                displayName: user.displayName,
              });
            }
          }
          resolve("Success");
        })
        .catch((err) => {
          console.error("TOT_TOTLeaderboard_getAllUsers \n", err);
          this.setState({
            showError: true,
            errorMessage:
              stringsConstants.TOTErrorMessage +
              " while getting users. Below are the details: \n" +
              JSON.stringify(err),
            showSuccess: false,
          });
          reject("Failed");
        });
    });
  }

  //get all user action for active tournament and bind to table
  private async getUserActions(): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        //get all users
        await this.getAllUsers();
        //get active tournament details
        let tournamentDetails =
          await commonService.getActiveTournamentDetails();
        if (tournamentDetails.length != 0) {
          this.setState({
            tournamentName: tournamentDetails[0]["Title"],
            tournamentDescription: tournamentDetails[0]["Description"],
          });
          //get active tournament's participants
          await commonService
            .getUserActions(this.state.tournamentName, allUsersDetails)
            .then((res) => {
              if (res.length > 0) {
                this.setState({
                  allUserActions: res,
                  isShowLoader: false,
                  userLoaded: "1",
                });
              } else if (res.length == 0) {
                this.setState({
                  userLoaded: "0",
                  showError: true,
                  noActiveParticipants: true,
                  errorMessage: stringsConstants.NoActiveParticipantsMessage,
                  isShowLoader: false,
                });
              } else if (res == "Failed") {
                console.error("TOT_TOTLeaderboard_getUserActions \n");
              }
            });
        } else {
          //no active tournaments
          this.setState({
            userLoaded: "1",
            showError: true,
            errorMessage: stringsConstants.NoActiveTournamentMessage,
            noActiveTournament: true,
            isShowLoader: false,
          });
        }
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
                />
                <span
                  className={styles.backLabel}
                  onClick={() => this.props.onClickCancel()}
                  title="Tournament of Teams"
                >
                  Tournament of Teams
                </span>
                <span className={styles.border}></span>
                <span className={styles.totLeaderboardLabel}>Leader Board</span>
              </div>              
              {this.state.tournamentName != "" && (
                <div className={styles.contentArea}>
                  {this.state.tournamentName != "" && (
                    <ul className={styles.listArea}>
                      <li className={styles.listVal}>
                        <span className={styles.labelHeading}>Tournament</span>:
                        <span className={styles.labelNormal}>
                          {this.state.tournamentName}
                        </span>
                      </li>
                      {this.state.tournamentDescription && (
                        <li className={styles.listVal}>
                          <span className={styles.labelHeading}>
                            Description
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
                        responsive
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
                          There are no active participants at the moment.
                          Be the first to participate and log an activity from&nbsp;
                          <span className={styles.myDashboardLink} 
                            onClick={() => this.props.onClickMyDashboardLink()}>
                            My Dashboard
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
                text="Back"
                title="Back"
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
