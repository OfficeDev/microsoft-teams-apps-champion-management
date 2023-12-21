import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import { Icon, initializeIcons } from "office-ui-fabric-react";
import * as React from "react";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import { RxJsEventEmitter } from "../events/RxJsEventEmitter";
import { EventData } from "../events/EventData";
import sideBarStyles from "../scss/TOTSideBar.module.scss";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

initializeIcons();

let commonService: commonServices;

export interface ITOTSideBarProps {
  context: any;
  siteUrl: string;
  currentUserDetails?: any;
  onClickCancel: Function;
}
interface ITOTSideBarState {
  showSuccess: Boolean;
  showError: Boolean;
  errorMessage: string;
  isShowLoader: boolean;
  userPoints: string;
  userRank: string;
  totalParticipants: string;
  userDisplayName: string;
  userEmail: string;
  allUserProps: any;
  isDesktop: boolean;
}
export default class TOTSideBar extends React.Component<
  ITOTSideBarProps,
  ITOTSideBarState
> {
  private readonly _eventEmitter: RxJsEventEmitter =
    RxJsEventEmitter.getInstance();
  constructor(props: ITOTSideBarProps, state: ITOTSideBarState) {
    super(props);
    //Set default values
    this.state = {
      showSuccess: false,
      showError: false,
      errorMessage: "",
      isShowLoader: false,
      userPoints: "",
      userRank: "",
      totalParticipants: "",
      userDisplayName: "",
      userEmail: "",
      allUserProps: [],
      isDesktop: true
    };
    //Create object for commonServices class
    commonService = new commonServices(this.props.context, this.props.siteUrl);
    this._eventEmitter.on(
      "rebindSideBar:start",
      this.getCurrentUserDetails.bind(this)
    );
  }
  //Get User Actions from list and bind to table
  public componentDidMount() {
    this.setState({
      isShowLoader: true,
    });
    //if props contains user details then set the state using the same, else re-evaluate and bind the user details
    if (
      this.props.currentUserDetails == undefined ||
      this.props.currentUserDetails.length == 0
    ) {
      this.getCurrentUserDetails();
    }
    else {
      let filterCurrentUser = this.props.currentUserDetails.filter(
        (e: any) =>
          e.Email === this.props.context.pageContext.user.email.toLowerCase()
      );
      if (filterCurrentUser.length > 0) {
        this.setState({
          isShowLoader: true,
          userDisplayName: this.props.context.pageContext.user.displayName,
          userEmail: filterCurrentUser[0].Email,
          userPoints: filterCurrentUser[0].Points,
          userRank: filterCurrentUser[0].Rank,
          totalParticipants: this.props.currentUserDetails.length,
        });
      } else {
        this.setState({
          userDisplayName: this.props.context.pageContext.user.displayName,
          userEmail: this.props.context.pageContext.user.email,
          userRank: "0",
          userPoints: "0",
          totalParticipants: this.props.currentUserDetails.length,
          isShowLoader: false,
        });
      }
    }

    // Adding window resize event listener while mounting the component
    window.addEventListener("resize", this.resize.bind(this));
    this.resize();
    //set css properties to Person card control
    this.updatePersonCardCSS();
  }

  // Set the state object for screen size
  resize = () => this.setState({ isDesktop: window.innerWidth > 568 });

  //set css properties to Person card control
  public updatePersonCardCSS() {
    setTimeout(() => {
      const sidebarPersonWrapper = document.getElementById("tot-sidebar-person-wrapper")?.querySelector("mgt-person")
        ?.shadowRoot?.querySelector("mgt-flyout")?.querySelector(".vertical");
      sidebarPersonWrapper?.setAttribute("style", "row-gap:10px;");
      if (this.state.isDesktop) {
        sidebarPersonWrapper?.querySelector(".avatar-wrapper")?.setAttribute("style", "width: 100px; height: 100px;");
        sidebarPersonWrapper?.querySelector(".details-wrapper")?.querySelector(".line1")?.setAttribute("style",
          "width: 180px;overflow-wrap: break-word;text-align: center;");
      }
      else {
        sidebarPersonWrapper?.querySelector(".avatar-wrapper")?.setAttribute("style", "");
        sidebarPersonWrapper?.querySelector(".details-wrapper")?.querySelector(".line1")
          ?.setAttribute("style", "font-size:14px;width: 140px;overflow-wrap: break-word;text-align: center;");
      }
    }, 5000);
  }

  public componentDidUpdate(prevProps: Readonly<ITOTSideBarProps>, prevState: Readonly<ITOTSideBarState>) {
    if (prevState.isDesktop !== this.state.isDesktop) {
      //set css properties to Person card control
      this.updatePersonCardCSS();
    }
  }

  private async getCurrentUserDetails(data?: EventData): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {

        let tournamentName: string;
        let allUserActions: any = [];

        //Setting varaible based on the value received from Event Emitter
        if (data != undefined)
          tournamentName = data.tournamentName;

        if (tournamentName != undefined && tournamentName != "") {
          //Getting user actions for the selected tournament to calculate points and rank
          await commonService.getUserActions(tournamentName).then((res) => {
            if (res.length > 0) {
              allUserActions = res;
              let filterCurrentUser = allUserActions.filter(
                (e: any) =>
                  e.Email ===
                  this.props.context.pageContext.user.email.toLowerCase()
              );
              //bind current user Ranks,Name,Points
              if (filterCurrentUser.length > 0) {
                this.setState({
                  userDisplayName: this.props.context.pageContext.user.displayName,
                  userEmail: filterCurrentUser[0].Email,
                  userRank: filterCurrentUser[0].Rank,
                  userPoints: filterCurrentUser[0].Points,
                  totalParticipants: res.length,
                  isShowLoader: false,
                });
              } else {
                this.setState({
                  userDisplayName: this.props.context.pageContext.user.displayName,
                  userEmail: this.props.context.pageContext.user.email,
                  userRank: "0",
                  userPoints: "0",
                  totalParticipants: res.length,
                  isShowLoader: false,
                });
              }
            } else if (res.length == 0) {
              //no active participants
              this.setState({
                userDisplayName: this.props.context.pageContext.user.displayName,
                userEmail: this.props.context.pageContext.user.email,
                userRank: "0",
                userPoints: "0",
                totalParticipants: "0",
                isShowLoader: false,
              });
            } else if (res == "Failed") {
              console.error("TOT_TOTSideBar_getUserActions \n");
            }
          });
        } else {
          //no active tournaments
          this.setState({
            userDisplayName: this.props.context.pageContext.user.displayName,
            userEmail: this.props.context.pageContext.user.email,
            userRank: "0",
            userPoints: "0",
            totalParticipants: "0",
            isShowLoader: false,
          });
        }
      } catch (error) {
        console.error("TOT_TOTSideBar_getUserActions \n", error);
        this.setState({
          showError: true,
          errorMessage:
            stringsConstants.TOTErrorMessage +
            " while getting user actions. Below are the details: \n" +
            JSON.stringify(error),
          showSuccess: false,
        });
      }
    });
  }
  public addDefaultSrc(ev: any) {
    ev.target.src = require("../assets/images/noprofile.png"); //if no profile then we are showing default image
  }
  public render() {
    return (
      <div className={sideBarStyles.totSideBar}>
        <div className={sideBarStyles.sideNav}>
          {this.state.userDisplayName != "" &&
            this.state.userDisplayName != undefined &&
            this.state.userEmail != undefined &&
            this.state.userEmail != "" && (
              <div className={sideBarStyles.imagePointsArea}>
                <div id="tot-sidebar-person-wrapper">
                  <Person
                    personQuery="me"
                    view={3}
                    personCardInteraction={1}
                    verticalLayout={true}
                    className={sideBarStyles.profileImageSideBar}
                  />
                </div>
                <div>
                  {/* here we are showing rank and points  */}
                  <div className={sideBarStyles.pointCircle}>
                    <div className={sideBarStyles.insideCircle}>
                      <div className={sideBarStyles.pointsScale}>
                        <div><Icon iconName="FavoriteStarFill" id="star" className={sideBarStyles.yellowStar} /></div>
                        <div className={sideBarStyles.pointsValueLabel}>{this.state.userPoints} {LocaleStrings.PointsLabel}</div>
                      </div>
                      <div className={sideBarStyles.line}></div>
                      <div className={sideBarStyles.globalRank}>
                        <div>{LocaleStrings.TournamentRankLabel}</div>
                        <div><span className={sideBarStyles.bold}>{this.state.userRank}</span> of {this.state.totalParticipants}</div>
                        <div>{LocaleStrings.ParticipantsLabel}</div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            )}
        </div>
      </div>
    );
  }
}
