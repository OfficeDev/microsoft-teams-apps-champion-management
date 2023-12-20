import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
//FluentUI controls
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Label } from "@fluentui/react/lib/Label";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Icon } from '@fluentui/react/lib/Icon';
import { ComboBox, IComboBox, IComboBoxOption } from '@fluentui/react/lib/ComboBox';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
//PNP
import { TreeView, ITreeItem, TreeViewSelectionMode, SelectChildrenMode } from "@pnp/spfx-controls-react/lib/TreeView";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTMyDashBoard.module.scss";
import TOTSidebar from "./TOTSideBar";
import { RxJsEventEmitter } from "../events/RxJsEventEmitter";
import { EventData } from "../events/EventData";
import * as LocaleStrings from 'ClbHomeWebPartStrings';

//Global Variables
let commonServiceManager: commonServices;
let currentUserEmail: string = "";
export interface ITOTMyDashboardProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
  currentThemeName?: string;
}
export interface ITOTMyDashboardState {
  actionsList: ITreeItem[];
  selectedActionsList: ITreeItem[];
  completedActionsList: ITreeItem[];
  showSuccess: boolean;
  showError: boolean;
  noActiveTournament: boolean;
  errorMessage: string;
  actionsError: boolean;
  tournamentName: any;
  showSpinner: boolean;
  noPendingActions: boolean;
  tournamentDescription: any;
  activeTournamentsList: Array<any>;
  myTournamentsList: Array<any>;
  activeTournamentName: any;
  myTournamentName: any;
  tournamentDescriptionList: Array<any>;
  treeViewSelectedKeys?: string[];
}

export default class TOTMyDashboard extends React.Component<ITOTMyDashboardProps, ITOTMyDashboardState> {
  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();
  public totMyDashboardTreeViewRef1: React.RefObject<HTMLDivElement>;
  public totMyDashboardTreeViewRef2: React.RefObject<HTMLDivElement>;

  constructor(props: ITOTMyDashboardProps, state: ITOTMyDashboardState) {
    super(props);

    //Set default values
    this.totMyDashboardTreeViewRef1 = React.createRef();
    this.totMyDashboardTreeViewRef2 = React.createRef();
    this.state = {
      actionsList: [],
      selectedActionsList: [],
      completedActionsList: [],
      showSuccess: false,
      showError: false,
      noActiveTournament: false,
      errorMessage: "",
      actionsError: false,
      tournamentName: "",
      showSpinner: false,
      noPendingActions: false,
      tournamentDescription: "",
      activeTournamentsList: [],
      myTournamentsList: [],
      activeTournamentName: "",
      myTournamentName: "",
      tournamentDescriptionList: [],
      treeViewSelectedKeys: [],

    };

    //Create object for CommonServices class
    commonServiceManager = new commonServices(this.props.context, this.props.siteUrl);

    //Bind Methods
    this.onActionSelected = this.onActionSelected.bind(this);
    this.getPendingActions = this.getPendingActions.bind(this);
    this.saveActions = this.saveActions.bind(this);
    this.getActiveTournamentActions = this.getActiveTournamentActions.bind(this);
    this.getMyTournamentActions = this.getMyTournamentActions.bind(this);
  }


  public componentDidMount() {
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
        await commonServiceManager.getActiveTournamentDetails();

      let activeTournamentsChoices: any = [];
      let myTournamentsChoices: any = [];
      let tournamentDescriptionChoices: any = [];
      //If active tournament found
      if (activeTournamentDetails.length > 0) {

        //Get current user's active tournament details
        let filterUserTournaments: string = "Title eq '" + currentUserEmail + "'";

        const currentUserTournaments: any[] =
          await commonServiceManager.getFilteredListItemsWithSpecificColumns(
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
          tournamentDescriptionList: tournamentDescriptionChoices
        });

        //When an user participates in an active tournament, move that tournament to My Tournaments dropdown.
        if (this.state.tournamentName != "") {
          this.setState({
            myTournamentName: this.state.tournamentName,
            activeTournamentName: null,
            tournamentName: this.state.tournamentName,
          });
          this.getPendingActions();
        }
        //Set the first option as a default tournament for My Tournaments dropdown
        else if (myTournamentsChoices.length > 0) {
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
          noActiveTournament: true
        });
      }
    }
    catch (error) {
      console.error("TOT_TOTMyDashboard_getActiveTournaments \n", error);
    }
  }

  //Set a value when an option is selected in My Tournaments dropdown and reset the Active Tournaments dropdown
  public getMyTournamentActions(ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void {
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

  //Refresh the tournament actions whenever the tournament name is selected
  public componentDidUpdate(prevProps: Readonly<ITOTMyDashboardProps>, prevState: Readonly<ITOTMyDashboardState>, snapshot?: any): void {
    try {
      if (prevState.tournamentName != this.state.tournamentName) {
        this.setState({ noPendingActions: false });
        if (this.state.tournamentName !== "")
          this.getPendingActions();
        //Refresh the points and rank in the sidebar when a tournament is selected in My Tournaments / Active tournaments dropdown
        this._eventEmitter.emit("rebindSideBar:start", {
          tournamentName: this.state.tournamentName,
        } as EventData);
      }

      //Update aria-label attribute to all Completed actions Treeview's info-icon-buttons.
      if (prevState.completedActionsList.length !== this.state.completedActionsList.length) {
        const completedActionsInfoButtons: any = this.totMyDashboardTreeViewRef2?.current?.getElementsByClassName('ms-Button--commandBar');
        for (let btn of completedActionsInfoButtons) {
          btn?.setAttribute("aria-label", "info");
        }

        //Update completed actions treeview expand buttons title attribute
        const treeElements: any = this.totMyDashboardTreeViewRef2?.current.querySelectorAll('div[class^="listItem_"]');
        for (let treeElement of treeElements) {
          const validBtnElement = treeElement?.querySelector(".ms-Button--icon");
          if (validBtnElement) {
            const exactValidBtnElement = treeElement?.querySelector('div[class^="itemContent_"]')?.querySelector('div[class^="labels_"]')?.textContent;
            validBtnElement?.setAttribute("title", exactValidBtnElement + " " + "Expanded");
          }
        }
      }

      //Update aria-label attribute to all Active actions Treeview's info-icon-buttons and Checkbox inputs.
      if (prevState.actionsList.length !== this.state.actionsList.length) {
        const pendingActionsInfoButtons: any = this.totMyDashboardTreeViewRef1?.current?.getElementsByClassName('ms-Button--commandBar');
        for (let btn of pendingActionsInfoButtons) {
          btn?.setAttribute("aria-label", "info");
        }

        //Update pending actions treeview expand buttons title attribute for accessibility
        const treeElements: any = this.totMyDashboardTreeViewRef1?.current.querySelectorAll('div[class^="listItem_"]');
        for (let treeElement of treeElements) {
          const validBtnElement = treeElement?.querySelector(".ms-Button--icon");
          if (validBtnElement) {
            const parentLabel = treeElement?.querySelector('div[class^="itemContent_"]')?.querySelector('div[class^="labels_"]')?.textContent;
            validBtnElement?.setAttribute("title", parentLabel + " " + "Expanded");
            const mainCheckbox = treeElement?.querySelector('div[class^="itemContent_"]')?.querySelector(".ms-Checkbox-label");
            mainCheckbox.setAttribute("aria-label", LocaleStrings.PendingActionsLabel + " " + parentLabel);
            const childElements = treeElement?.nextElementSibling?.querySelectorAll("div[class^='itemContent_'");
            for (let childElement of childElements) {
              const childLabel = childElement?.querySelector("div[class^='labels_']")?.childNodes[0]?.textContent;
              const childCheckbox = childElement?.querySelector(".ms-Checkbox-label");
              childCheckbox.setAttribute("aria-label", parentLabel + " " + childLabel);
            }
          }
        }
      }
    }
    catch (error: any) {
      console.error("CMP_TOT_TOTMyDashboard_componentDidUpdate \n", error);
    }
  }

  //On select of a tree node change the state of selected actions
  private onActionSelected(items: any) {
    this.setState({ selectedActionsList: items, treeViewSelectedKeys: items["key"] });
  }

  //Get Actions from Tournament Actions list and bind it to Treeview
  private async getPendingActions() {
    console.log(stringsConstants.TotLog + "Getting actions from Tournament Actions list.");
    try {
      // Reset state variables
      this.setState({
        actionsList: [],
        completedActionsList: [],
        selectedActionsList: [],
        actionsError: false,
        treeViewSelectedKeys: [],
      });


      //Get current users's email
      currentUserEmail =
        this.props.context.pageContext.user.email.toLowerCase();

      let filterActive: string =
        "Title eq '" +
        this.state.tournamentName.replace(/'/g, "''") +
        "'";
      let filterUserTournaments: string =
        "Tournament_x0020_Name eq '" +
        this.state.tournamentName.replace(/'/g, "''") +
        "'" +
        " and Title eq '" +
        currentUserEmail +
        "'";

      //Set the description for selected tournament
      let tournmentDesc = this.state.tournamentDescriptionList.find((item) => item.key == this.state.tournamentName);

      this.setState({
        tournamentDescription: tournmentDesc.text
      });

      //Get all actions for the tournament from "Tournament Actions" list
      const allTournamentsActionsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentActionsMasterList,
          filterActive
        );

      //Sort on Category
      allTournamentsActionsArray.sort((a, b) => a.Category.localeCompare(b.Category));

      //Get all actions completed by the current user for the current tournament
      const userActionsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.UserActionsList,
          filterUserTournaments
        );

      let treeItemsArray: ITreeItem[] = [];
      let completedTreeItemsArray: ITreeItem[] = [];

      //Build the Parent Nodes(Categories) in Treeview. Skip the items which are already completed by the user in "User Actions" list
      allTournamentsActionsArray.forEach((vAction) => {
        //Check if the category is present in the 'User Actions' list
        let compareCategoriesArray = userActionsArray.filter((elArray) => {
          return (
            elArray.Action == vAction["Action"] &&
            elArray.Category == vAction["Category"]
          );
        });
        const tree: ITreeItem = {
          key: vAction["Category"],
          label: vAction["Category"],
          children: [],
        };

        //If the category is not present in User Actions list add it to 'Pending Tree view'
        let found: boolean;
        if (compareCategoriesArray.length == 0) {
          //Check if Category is already added to the Treeview. If yes, skip adding.
          found = treeItemsArray.some((value) => {
            return value.label === vAction["Category"];
          });
          if (!found) treeItemsArray.push(tree);
        }
        //If the category is present in User Actions list add it to 'Completed Tree view'
        else {
          //Check if Category is already added to the Treeview. If yes, skip adding.
          found = completedTreeItemsArray.some((value) => {
            return value.label === vAction["Category"];
          });
          if (!found) completedTreeItemsArray.push(tree);
        }
      }); //For Loop

      //Build the child nodes(Actions) in Treeview. Skip the items which are already completed by the user in "User Actions" list
      allTournamentsActionsArray.forEach((vAction) => {
        //Check if the action is present in the 'User Actions' list
        let compareActionsArray = userActionsArray.filter((elChildArray) => {
          return (
            elChildArray.Action == vAction["Action"] &&
            elChildArray.Category == vAction["Category"]
          );
        });

        //If the action is  not present in User Actions list add it to 'Pending Tree view'
        let tree: ITreeItem;
        if (compareActionsArray.length == 0) {
          if (vAction["HelpURL"] === 'null' || vAction["HelpURL"] == "") {
            tree = {
              key: vAction.Id,
              label: vAction["Action"],
              data:
                vAction["Category"] +
                stringsConstants.StringSeperator +
                vAction["HelpURL"],
              subLabel:
                vAction["Points"] +
                stringsConstants.PointsDisplayString +
                vAction["Description"]
            };
          }
          else {
            tree = {
              key: vAction.Id,
              label: vAction["Action"],
              data:
                vAction["Category"] +
                stringsConstants.StringSeperator +
                vAction["HelpURL"],
              subLabel:
                vAction["Points"] +
                stringsConstants.PointsDisplayString +
                vAction["Description"],
              actions: [
                {
                  iconProps: {
                    iconName: "Info",
                    title: LocaleStrings.MyDashboardInfoIconMessage
                  },
                  id: "GetItem",
                  actionCallback: (treeItem: ITreeItem) => {
                    window.open(vAction["HelpURL"]);
                  },
                },
              ],
            };
          }
          let treeCol: Array<ITreeItem> = treeItemsArray.filter((value) => {
            return value.label == vAction["Category"];
          });
          if (treeCol.length != 0) {
            treeCol[0].children.push(tree);
          }
        }
        //If the action present in User Actions list add it to 'Completed Tree view'
        else {
          if (vAction["HelpURL"] === 'null' || vAction["HelpURL"] == "") {
            tree = {
              key: vAction.Id,
              label: vAction["Action"],
              data:
                vAction["Category"] +
                stringsConstants.StringSeperator +
                vAction["HelpURL"],
              subLabel:
                vAction["Points"] +
                stringsConstants.PointsDisplayString +
                vAction["Description"],
              iconProps: {
                iconName: "SkypeCheck",
              },
            };
          }
          else {
            tree = {
              key: vAction.Id,
              label: vAction["Action"],
              data:
                vAction["Category"] +
                stringsConstants.StringSeperator +
                vAction["HelpURL"],
              subLabel:
                vAction["Points"] +
                stringsConstants.PointsDisplayString +
                vAction["Description"],
              iconProps: {
                iconName: "SkypeCheck",
              },
              actions: [
                {
                  iconProps: {
                    iconName: "Info",
                    title: LocaleStrings.MyDashboardInfoIconMessage
                  },
                  id: "GetItem",
                  actionCallback: (treeItem: ITreeItem) => {
                    window.open(vAction["HelpURL"]);
                  },
                },
              ],
            };
          }
          let treeColCompleted: Array<ITreeItem> =
            completedTreeItemsArray.filter((value) => {
              return value.label == vAction["Category"];
            });
          if (treeColCompleted.length != 0) {
            treeColCompleted[0].children.push(tree);
          }
        }
      }); //For loop

      if (treeItemsArray.length == 0)
        this.setState({ noPendingActions: true });
      this.setState({
        actionsList: treeItemsArray,
        completedActionsList: completedTreeItemsArray,
      });

    } catch (error) {
      console.error("TOT_TOTMyDashboard_getPendingActions \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          " while retrieving actions list. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  // Save Tournament Details to SP Lists 'Actions List' and 'Tournament Actions'
  private async saveActions() {
    try {
      this.setState({ showSpinner: true, actionsError: false });
      if (this.state.selectedActionsList.length > 0) {
        let selectedTreeArray: ITreeItem[] = this.state.selectedActionsList;
        //Loop through actions selected and create a list item for each treeview selection
        let createActionsPromise: any = [];
        let checkActionsPromise: any = [];
        for (let item of selectedTreeArray) {
          //Skip parent node for treeview which is not an action
          if (item.data != undefined) {
            //Insert User Action only if its not already there.
            let filterUserTournaments: string =
              "Tournament_x0020_Name eq '" +
              this.state.tournamentName.replace(/'/g, "''") +
              "'" +
              " and Title eq '" +
              currentUserEmail +
              "'" +
              " and Action eq '" +
              item.label.replace(/'/g, "''") +
              "'";
            let checkActionsPresent =
              await commonServiceManager.getItemsWithOnlyFilter(
                stringsConstants.UserActionsList,
                filterUserTournaments
              );
            checkActionsPromise.push(checkActionsPresent);
          }
        }
        Promise.all(checkActionsPromise).then(async (responseObj) => {
          let filterChildNodesArray = selectedTreeArray.filter(
            (eFilter) => eFilter.data != undefined
          );
          for (let iCount = 0; iCount < responseObj.length; iCount++) {
            if (responseObj[iCount].length == 0) {
              if (filterChildNodesArray[iCount].data != undefined) {
                let submitObject: any = {
                  Title: currentUserEmail,
                  Tournament_x0020_Name: this.state.tournamentName,
                  Action: filterChildNodesArray[iCount].label,
                  Category: filterChildNodesArray[iCount].data.split(
                    stringsConstants.StringSeperator
                  )[0],
                  Points: filterChildNodesArray[iCount].subLabel
                    .split(stringsConstants.StringSeperatorPoints)[0]
                    .replace(stringsConstants.PointsReplaceString, ""),
                  UserName: this.props.context.pageContext.user.displayName
                };
                let createItems = await commonServiceManager.createListItem(
                  stringsConstants.UserActionsList,
                  submitObject
                );
                createActionsPromise.push(createItems);
              }
            }
          }
          this.setState({ actionsList: [], selectedActionsList: [] });
          Promise.all(createActionsPromise).then(() => {
            this.getActiveTournaments().then(() => {
              this.setState({ showSpinner: false });
              this._eventEmitter.emit("rebindSideBar:start", {
                tournamentName: this.state.tournamentName,
              } as EventData);
            });
          });
        });
      }
      //No Action selected in Treeview
      else {
        this.setState({ actionsError: true, showSpinner: false });
      }
    } catch (error) {
      console.error("TOT_TOTMyDashboard_saveActions \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          " while saving your actions. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

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

  //Render Method
  public render(): React.ReactElement<ITOTMyDashboardProps> {
    const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
    return (
      <div className={`${styles.container}${isDarkOrContrastTheme ? " " + styles.containerDarkContrast : ""}`}>
        <div className={styles.totDashboardContent}>
          <TOTSidebar
            siteUrl={this.props.siteUrl}
            context={this.props.context}
            onClickCancel={() => this.props.onClickCancel()}
          />
          <div className={styles.contentTab}>
            <div>
              <div className={styles.totDashboardPath}>
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
                <span className={styles.border} aria-live="polite" role="alert" aria-label={LocaleStrings.TOTMyDashboardPageTitle + " Page"} />
                <span className={styles.totDashboardLabel}>{LocaleStrings.TOTMyDashboardPageTitle}</span>
              </div>
              {this.state.showError && (
                <div>
                  {this.state.noActiveTournament ? (
                    <div>
                      <Label className={styles.noTourErrorMessage}>{this.state.errorMessage}</Label>
                      <DefaultButton
                        text={LocaleStrings.BackButton}
                        title={LocaleStrings.BackButton}
                        iconProps={{ iconName: 'NavigateBack' }}
                        onClick={() => this.props.onClickCancel()}
                        className={styles.totMyDashboardBackBtn}
                      />
                    </div>
                  )
                    :
                    <Label className={styles.errorMessage}>{this.state.errorMessage}</Label>
                  }
                </div>
              )}
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
                        id="tot-my-dashboard-combobox-info"
                      >
                        <Icon
                          aria-label="Info"
                          aria-describedby="tot-my-dashboard-combobox-info"
                          iconName="Info"
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
                        className: `totMdbComboCallout${isDarkOrContrastTheme ? ' totMdbComboCallout--' + this.props.currentThemeName : ""}`,
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
                      ariaLabel={LocaleStrings.ActiveTournamentLabel + ' list'}
                      useComboBoxAsMenuWidth={true}
                      calloutProps={{
                        className: `totMdbComboCallout${isDarkOrContrastTheme ? ' totMdbComboCallout--' + this.props.currentThemeName : ""}`,
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
                            <span className={styles.labelHeading + " " + styles.descriptionHeading}>{LocaleStrings.DescriptionLabel}</span>
                            <span className={styles.descriptionColon}>:</span>
                            <span className={styles.labelNormal}>{this.state.tournamentDescription}</span>
                          </li>
                        )}
                      </ul>
                    )}
                  </div>
                </Col>
              </Row>
            )}
            {this.state.tournamentName != "" && (
              <div className={styles.contentArea}>
                <Row xl={2} lg={2} md={2} sm={1} xs={1}>
                  <Col xl={6} lg={6} md={6} sm={12} xs={12}>
                    <h2 tabIndex={0} role="heading">
                      <Label className={styles.subHeaderUnderline}>{LocaleStrings.PendingActionsLabel}</Label>
                    </h2>
                    {this.state.noPendingActions && (
                      <Label className={styles.successMessage}>
                        <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
                        {LocaleStrings.PendingActionsSuccessMessage}
                      </Label>
                    )}
                    {this.state.actionsList !== undefined && (
                      <div className={styles.myDashBoardTreeView1} ref={this.totMyDashboardTreeViewRef1}>
                        {this.state.actionsList.length > 0 &&
                          <TreeView
                            items={this.state.actionsList}
                            defaultExpanded={true}
                            selectionMode={TreeViewSelectionMode.Multiple}
                            selectChildrenMode={SelectChildrenMode.Select | SelectChildrenMode.Unselect}
                            showCheckboxes={true}
                            defaultSelectedKeys={this.state.treeViewSelectedKeys}
                            onSelect={this.onActionSelected}
                            onExpandCollapse={(item, isExpanded: boolean) => {
                              //Get all the tree structured div elements from the this.totMyDashboardTreeViewRef1 (DOM)
                              const treeElements: any = this.totMyDashboardTreeViewRef1?.current.querySelectorAll('div[class^="listItem_"]');
                              for (let treeElement of treeElements) {
                                //Get Expand/collapse icon button elements from  each tree element
                                const validBtnElement = treeElement?.querySelector(".ms-Button--icon");
                                if (validBtnElement) {
                                  //Get current expand/collapse button from selected tree element
                                  const exactValidBtnElement = treeElement?.querySelector('div[class^="itemContent_"]')?.querySelector('div[class^="labels_"]')?.textContent;
                                  if (exactValidBtnElement.trim() === item.label) {
                                    //Update Title attribute
                                    if (isExpanded) {
                                      setTimeout(() => {
                                        validBtnElement?.setAttribute("title", item.label + " " + "Expanded");
                                        const childElements = treeElement?.nextElementSibling?.querySelectorAll("div[class^='itemContent_'");
                                        for (let childElement of childElements) {
                                          const childLabel = childElement?.querySelector("div[class^='labels_']")?.childNodes[0].textContent;
                                          const childCheckbox = childElement?.querySelector(".ms-Checkbox-label");
                                          childCheckbox.setAttribute("aria-label", item.label + " " + childLabel);
                                        }
                                      }, 5);
                                    }
                                    else {
                                      setTimeout(() => {
                                        validBtnElement?.setAttribute("title", item.label + " " + "Collapsed");
                                      }, 5);
                                    }
                                    break;
                                  }
                                }
                              }
                            }}
                          />
                        }
                      </div>
                    )}
                    {this.state.actionsError && (
                      <Label className={styles.errorMessage}>{LocaleStrings.SelectActionsErrorMessage}</Label>
                    )}
                    {this.state.showSpinner && (
                      <Spinner label={LocaleStrings.FormSavingMessage} size={SpinnerSize.large} />
                    )}
                    <div className={styles.btnArea}>
                      {this.state.actionsList.length != 0 && (
                        <PrimaryButton
                          text={LocaleStrings.SaveButton}
                          title={LocaleStrings.SaveButton}
                          iconProps={{ iconName: 'Save' }}
                          onClick={this.saveActions}
                          className={styles.saveBtn}
                        />
                      )}
                      &nbsp; &nbsp;
                      <PrimaryButton
                        text={LocaleStrings.BackButton}
                        title={LocaleStrings.BackButton}
                        iconProps={{ iconName: 'NavigateBack' }}
                        onClick={() => this.props.onClickCancel()}
                        className={styles.backBtnToggle1 + " " + styles.totMyDashboardBackBtn}
                      />
                    </div>
                  </Col>
                  <Col xl={6} lg={6} md={6} sm={12} xs={12}>
                    <h2 tabIndex={0} role="heading">
                      <Label className={styles.subHeaderUnderline}>{LocaleStrings.CompletedActionsLabel}</Label>
                    </h2>
                    {this.state.completedActionsList !== undefined && (
                      <div className={styles.myDashBoardTreeView2} ref={this.totMyDashboardTreeViewRef2}>
                        {this.state.completedActionsList.length > 0 &&
                          <TreeView
                            items={this.state.completedActionsList}
                            defaultExpanded={true}
                            onExpandCollapse={(item, isExpanded: boolean) => {
                              //Get all the tree structured div elements from the this.totMyDashboardTreeViewRef2(DOM)
                              const treeElements: any = this.totMyDashboardTreeViewRef2?.current.querySelectorAll('div[class^="listItem_"]');
                              for (let treeElement of treeElements) {
                                //Get Expand/collapse icon button elements from  each tree element
                                const validBtnElement = treeElement?.querySelector(".ms-Button--icon");
                                if (validBtnElement) {
                                  //Get current expand/collapse button from selected tree element
                                  const exactValidBtnElement = treeElement?.querySelector('div[class^="itemContent_"]')?.querySelector('div[class^="labels_"]')?.textContent;
                                  if (exactValidBtnElement.trim() === item.label) {
                                    //Update Title attribute
                                    if (isExpanded) {
                                      setTimeout(() => {
                                        validBtnElement?.setAttribute("title", item.label + " " + "Expanded");
                                      }, 5);
                                    }
                                    else {
                                      setTimeout(() => {
                                        validBtnElement?.setAttribute("title", item.label + " " + "Collapsed");
                                      }, 5);
                                    }
                                    break;
                                  }
                                }
                              }
                            }}
                          />
                        }
                      </div>
                    )}
                  </Col>
                  <Col xl={12} lg={12} md={12} sm={12} xs={12}>
                    <PrimaryButton
                      text={LocaleStrings.BackButton}
                      title={LocaleStrings.BackButton}
                      iconProps={{ iconName: 'NavigateBack' }}
                      onClick={() => this.props.onClickCancel()}
                      className={styles.backBtnToggle2 + " " + styles.totMyDashboardBackBtn}
                    />
                  </Col>
                </Row>
              </div>
            )}
          </div>
        </div>
      </div> //Final DIV
    );
  }
}
