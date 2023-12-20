import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
//React Boot Strap
import BootstrapTable from "react-bootstrap-table-next";
import paginationFactory from "react-bootstrap-table2-paginator";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
//FluentUI controls
import { DefaultButton } from "@fluentui/react";
import { Icon } from '@fluentui/react/lib/Icon';
import { Label } from "@fluentui/react/lib/Label";
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTLeaderBoard.module.scss";
import TOTSidebar from "./TOTSideBar";
import { RxJsEventEmitter } from "../events/RxJsEventEmitter";
import { EventData } from "../events/EventData";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";

//Global variables
let commonService: commonServices;
let currentUserEmail: string = "";

export interface ITOTLeaderBoardProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
  onClickMyDashboardLink: Function;
  currentThemeName?: string;
}
interface ITOTLeaderBoardState {
  showSuccess: boolean;
  showError: boolean;
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

export default class TOTLeaderBoard extends React.Component<ITOTLeaderBoardProps, ITOTLeaderBoardState> {
  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  private leaderboardTableRef: React.RefObject<HTMLDivElement>;

  constructor(props: ITOTLeaderBoardProps, state: ITOTLeaderBoardState) {
    super(props);

    //Create ref for leaderboard table
    this.leaderboardTableRef = React.createRef();

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
      userLoaded: ""
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

      let activeTournamentsChoices: any = [];
      let myTournamentsChoices: any = [];
      let tournamentDescriptionChoices: any = [];
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
        activeTournamentDetails.forEach((eachTournament) => {

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
        activeTournamentsChoices.sort((a: any, b: any) => a.text.localeCompare(b.text));
        myTournamentsChoices.sort((a: any, b: any) => a.text.localeCompare(b.text));

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

    //Accessibility
    //Update Page Per Size Dropdown and its button to be accessibile through Keyboard
    if (prevState.allUserActions !== this.state.allUserActions && this.state.allUserActions.length > 0) {
      //Outside Click event for page dropdown button
      document.addEventListener("click", this.onPageDropdownBtnOutsideClick);

      //Set aria-label to leaderboard table element
      const tableElement = this.leaderboardTableRef?.current?.querySelector("table");
      tableElement.setAttribute("aria-label", this.state.tournamentName + " " + LocaleStrings.TOTLeaderBoardPageTitle);

      //Get Page Dropdown button element
      const sizePerPageBtnElement: any = this.leaderboardTableRef?.current?.querySelector("#pageDropDown");
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
      const paginationFirstBtn = this.leaderboardTableRef?.current?.querySelector(".pagination").getElementsByTagName('a')[0];
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
  }

  //Remove Document click event listener on Unmount of Component
  public componentWillUnmount(): void {
    document.removeEventListener("click", this.onPageDropdownBtnOutsideClick);
  }

  //Get Page Dropdown Button's Callout Element from DOM
  public getUlElement = () => {
    const ulElements: any = this.leaderboardTableRef?.current?.getElementsByTagName("ul");
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

  //Get all user action for active tournament and bind to table
  private async getUserActions(): Promise<any> {
    try {
      //Set the description for selected tournament
      let tournmentDesc = this.state.tournamentDescriptionList.find((item) => item.key == this.state.tournamentName);
      this.setState({ tournamentDescription: tournmentDesc.text });
      //get active tournament's participants
      await commonService.getUserActions(this.state.tournamentName)
        .then((res) => {
          if (res.length > 0) {
            this.setState({
              allUserActions: res,
              isShowLoader: false,
              userLoaded: "1",
            });
          }
          else if (res.length == 0) {
            this.setState({
              allUserActions: [],
              userLoaded: "0",
              showError: true,
              noActiveParticipants: true,
              errorMessage: LocaleStrings.NoActiveParticipantsMessage,
              isShowLoader: false,
            });
          }
          else if (res == "Failed") {
            console.error("TOT_TOTLeaderboard_getUserActions \n");
          }
        });
    }
    catch (error) {
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
  }

  //Get Pagination Properties
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
              const sizePerPageBtnElement: HTMLButtonElement = this.leaderboardTableRef?.current?.querySelector("#pageDropDown");
              sizePerPageBtnElement.setAttribute("aria-label", stringsConstants.sizePerPageLabel + " " + options.text);
              sizePerPageBtnElement?.focus();
            }}
            onKeyDown={(evt: any) => {
              const sizePerPageUlElement = this.getUlElement();
              const sizePerPageBtnElement: HTMLButtonElement = this.leaderboardTableRef?.current?.querySelector("#pageDropDown");
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
      const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} results` : "";
      return (<span className="react-bootstrap-table-pagination-total" aria-live="polite" role="status">
        &nbsp;{resultsFound}
      </span>
      );
    }
  });

  //On menu open add the attributes to fix the position issue in IOS for Accessibility
  private onMenuOpen = (listboxId: string) => {
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
  participantFormatter = (cell: any) => {
    return (
      <Person
        personQuery={cell}
        view={3}
        personCardInteraction={1}
        className="particpant-person-card"
      />
    );
  }

  public render(): React.ReactElement<ITOTLeaderBoardProps> {
    const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
    const columns = [
      {
        dataField: "Rank",
        text: LocaleStrings.RankLabel,
      },
      {
        dataField: "Email",
        text: LocaleStrings.UserLabel,
        formatter: this.participantFormatter
      },
      {
        dataField: "Points",
        text: LocaleStrings.PointsLabel,
      },
    ];
    return (
      <div className={`${styles.totLeaderboardWrapper}${isDarkOrContrastTheme ? " " + styles.totLeaderboardWrapperDarkContrast : ""}`}>
        {this.state.isShowLoader && <div className={styles.load} aria-label="Loading Leaderboard Page" aria-live="polite" role="alert"></div>}
        <div className={styles.container}>
          <div className={styles.totLeaderboardContent}>
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
                <span className={styles.border} />
                <span className={styles.totLeaderboardLabel} aria-live="polite" role="alert" tabIndex={0}>{LocaleStrings.TOTLeaderBoardPageTitle}</span>
              </div>
              <div className={styles.dropdownArea}>
                <Row xl={3} lg={3} md={3} sm={1} xs={1}>
                  {this.state.myTournamentsList.length > 0 && (
                    <Col xl={5} lg={5} md={5} sm={12} xs={12}>
                      <span className={styles.labelHeading}>{LocaleStrings.MyTournamentsLabel} :
                        <TooltipHost
                          content={LocaleStrings.MyTournamentsTooltip}
                          calloutProps={{ gapSpace: 0 }}
                          hostClassName={styles.tooltipHostStyles}
                          delay={window.innerWidth < stringsConstants.MobileWidth ? 0 : 2}
                          id="tot-leaderboard-combobox-info"
                        >
                          <Icon
                            aria-label="Info"
                            iconName="Info"
                            aria-describedby="tot-leaderboard-combobox-info"
                            className={styles.myTournamentInfoIcon}
                            tabIndex={0}
                            role="button"
                          />
                        </TooltipHost>
                      </span>
                      <ComboBox className={styles.dropdownCol}
                        placeholder={LocaleStrings.SelectTournamentPlaceHolder}
                        selectedKey={this.state.myTournamentName}
                        options={this.state.myTournamentsList}
                        onChange={this.getMyTournamentActions.bind(this)}
                        ariaLabel={LocaleStrings.MyTournamentsLabel + ' list'}
                        useComboBoxAsMenuWidth={true}
                        calloutProps={{
                          className: `totLbComboCallout${isDarkOrContrastTheme ? ' totLbComboCallout--' + this.props.currentThemeName : ""}`,
                          directionalHintFixed: true, doNotLayer: true
                        }}
                        allowFreeInput={true}
                        persistMenu={true}
                        id="my-tournaments-listbox"
                        onMenuOpen={() => this.onMenuOpen("my-tournaments-listbox")}
                      />
                    </Col>
                  )}
                  {this.state.myTournamentsList.length > 0 && this.state.activeTournamentsList.length > 0 && (
                    <Col xl={2} lg={2} md={2} sm={12} xs={12} className={styles.labelCol} >
                      <span className={styles.labelHeading}>{LocaleStrings.OrLabel}</span>
                    </Col>
                  )}
                  {this.state.activeTournamentsList.length > 0 && (
                    <Col xl={5} lg={5} md={5} sm={12} xs={12}>
                      <span className={styles.labelHeading}>{LocaleStrings.ActiveTournamentLabel} : </span>
                      <ComboBox className={styles.dropdownCol}
                        placeholder={LocaleStrings.SelectTournamentPlaceHolder}
                        selectedKey={this.state.activeTournamentName}
                        options={this.state.activeTournamentsList}
                        onChange={this.getActiveTournamentActions.bind(this)}
                        ariaLabel={LocaleStrings.ActiveTournamentLabel + " list"}
                        useComboBoxAsMenuWidth={true}
                        calloutProps={{
                          className: `totLbComboCallout${isDarkOrContrastTheme ? ' totLbComboCallout--' + this.props.currentThemeName : ""}`,
                          directionalHintFixed: true, doNotLayer: true
                        }}
                        allowFreeInput={true}
                        persistMenu={true}
                        id="active-tournaments-listbox"
                        onMenuOpen={() => this.onMenuOpen("active-tournaments-listbox")}
                      />
                    </Col>
                  )}
                </Row>
              </div>
              {this.state.tournamentName != "" && (
                <Row xl={1} lg={1} md={1} sm={1} xs={1}>
                  <Col xl={12} lg={12} md={12} sm={12} xs={12}>
                    <div>
                      {this.state.tournamentName != "" && (
                        <ul className={styles.listArea}>
                          {this.state.tournamentDescription && (
                            <li className={styles.listVal}>
                              <span className={styles.labelHeading + " " + styles.descriptionHeading}>
                                {LocaleStrings.DescriptionLabel}
                              </span>
                              <span className={styles.descriptionColon}>:</span>
                              <span className={styles.labelNormal}>
                                {this.state.tournamentDescription}
                              </span>
                            </li>
                          )}
                        </ul>
                      )}
                      <div className={styles.table} ref={this.leaderboardTableRef}>
                        {this.state.allUserActions.length > 0 ? (
                          <BootstrapTable
                            table-responsive
                            bordered
                            hover
                            keyField="Rank"
                            data={this.state.allUserActions}
                            columns={columns}
                            pagination={this.pagination}
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
                  </Col>
                </Row>
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
                  iconProps={{ iconName: 'NavigateBack' }}
                  onClick={() => this.props.onClickCancel()}
                  className={styles.totLeaderboardBackBtn}
                  ariaHidden={this.state.isShowLoader}
                />
              </div>
            </div>
          </div>
        </div>
      </div> //outer div
    ); //close return
  } //end of render
}
