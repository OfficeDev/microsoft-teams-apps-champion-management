import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTMyDashBoard.module.scss";
import TOTSidebar from "./TOTSideBar";
import { RxJsEventEmitter } from "../events/RxJsEventEmitter";

//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";

//FluentUI controls
import { IButtonStyles, PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import { Label } from "@fluentui/react/lib/Label";
import { Spinner, SpinnerSize } from "@fluentui/react/lib/Spinner";
import { Icon, IIconProps } from '@fluentui/react/lib/Icon';

//PNP
import {
  TreeView,
  ITreeItem,
  TreeViewSelectionMode,
} from "@pnp/spfx-controls-react/lib/TreeView";

//Global Variables
let commonServiceManager: commonServices;
let currentUserEmail: string = "";
export interface ITOTMyDashboardProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
}

const backBtnStyles: Partial<IButtonStyles> = {
  root: {
    marginLeft: "1.5%",
    marginTop: "1.5%",
    borderColor: "#33344A",
    backgroundColor: "white",
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

const saveIcon: IIconProps = { iconName: 'Save' };
const backIcon: IIconProps = { iconName: 'NavigateBack' };

interface ITOTMyDashboardState {
  actionsList: ITreeItem[];
  selectedActionsList: ITreeItem[];
  completedActionsList: ITreeItem[];
  showSuccess: boolean;
  showError: boolean;
  noActiveTournament: boolean;
  errorMessage: string;
  actionsError: boolean;
  tournamentName: string;
  showSpinner: boolean;
  noPendingActions: boolean;
  tournamentDescription: string;
}

export default class TOTMyDashboard extends React.Component<
  ITOTMyDashboardProps,
  ITOTMyDashboardState
> {
  private readonly _eventEmitter: RxJsEventEmitter =
    RxJsEventEmitter.getInstance();
  constructor(props: ITOTMyDashboardProps, state: ITOTMyDashboardState) {
    super(props);
    //Set default values
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
    };
    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );

    //Bind Methods
    this.onActionSelected = this.onActionSelected.bind(this);
    this.getPendingActions = this.getPendingActions.bind(this);
    this.saveActions = this.saveActions.bind(this);
  }

  //Get Actions from Master list and bind it to treeview on app load
  public componentDidMount() {
    //Get Actions from Master list and bind it to Treeview
    this.getPendingActions();
  }

  //On select of a tree node change the state of selected actions
  private onActionSelected(items: ITreeItem[]) {
    this.setState({ selectedActionsList: items });
  }

  //Get Actions from Master list and bind it to Treeview
  private async getPendingActions() {
    console.log(stringsConstants.TotLog + "Getting actions from master list.");
    try {
      //Get current users's email
      currentUserEmail =
        this.props.context.pageContext.user.email.toLowerCase();

      //Get current active tournament details
      let tournamentDetails: any[] =
        await commonServiceManager.getActiveTournamentDetails();

      //If active tournament found
      if (tournamentDetails.length != 0) {
        this.setState({
          tournamentName: tournamentDetails[0]["Title"],
          tournamentDescription: tournamentDetails[0]["Description"],
        });
        let filterActive: string =
          "Title eq '" +
          tournamentDetails[0]["Title"].replace(/'/g, "''") +
          "'";
        let filterUserTournaments: string =
          "Tournament_x0020_Name eq '" +
          tournamentDetails[0]["Title"].replace(/'/g, "''") +
          "'" +
          " and Title eq '" +
          currentUserEmail +
          "'";
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
        var treeItemsArray: ITreeItem[] = [];
        var completedTreeItemsArray: ITreeItem[] = [];

        //Build the Parent Nodes(Categories) in Treeview. Skip the items which are already completed by the user in "User Actions" list
        await allTournamentsActionsArray.forEach((vAction) => {
          //Check if the category is present in the 'User Actions' list
          var compareCategoriesArray = userActionsArray.filter((elArray) => {
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
          var found: boolean;
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
        await allTournamentsActionsArray.forEach((vAction) => {
          //Check if the action is present in the 'User Actions' list
          var compareActionsArray = userActionsArray.filter((elChildArray) => {
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
                      title: "Find out more about this action"
                    },
                    id: "GetItem",
                    actionCallback: async (treeItem: ITreeItem) => {
                      window.open(vAction["HelpURL"]);
                    },
                  },
                ],
              };
            }
            var treeCol: Array<ITreeItem> = treeItemsArray.filter((value) => {
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
                      title: "Find out more about this action"
                    },
                    id: "GetItem",
                    actionCallback: async (treeItem: ITreeItem) => {
                      window.open(vAction["HelpURL"]);
                    },
                  },
                ],
              };
            }
            var treeColCompleted: Array<ITreeItem> =
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
      } // IF END

      //If there is no active tournament
      else {
        this.setState({
          showError: true,
          errorMessage: stringsConstants.NoActiveTournamentMessage,
          noActiveTournament: true
        });
      }
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
        var selectedTreeArray: ITreeItem[] = this.state.selectedActionsList;
        //Loop through actions selected and create a list item for each treeview selection
        let createActionsPromise = [];
        let checkActionsPromise = [];
        for (let counter = 0; counter < selectedTreeArray.length; counter++) {
          //Skip parent node for treeview which is not an action
          if (selectedTreeArray[counter].data != undefined) {
            //Insert User Action only if its not already there.
            let filterUserTournaments: string =
              "Tournament_x0020_Name eq '" +
              this.state.tournamentName.replace(/'/g, "''") +
              "'" +
              " and Title eq '" +
              currentUserEmail +
              "'" +
              " and Action eq '" +
              selectedTreeArray[counter].label +
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
            this.getPendingActions().then(() => {
              this.setState({ showSpinner: false });
              this._eventEmitter.emit("rebindSideBar:start", {
                currentNumber: "1",
              });
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

  //Render Method
  public render(): React.ReactElement<ITOTMyDashboardProps> {
    return (
      <div className={styles.container}>
        <div className={styles.totSideBar}>
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
                />
                <span
                  className={styles.backLabel}
                  onClick={() => this.props.onClickCancel()}
                  title="Tournament of Teams"
                >
                  Tournament of Teams
                </span>
                <span className={styles.border}></span>
                <span className={styles.totDashboardLabel}>My Dashboard</span>
              </div>
              {this.state.showError && (
                <div>
                  {this.state.noActiveTournament ? (
                    <div>
                      <Label className={styles.noTourErrorMessage}>
                        {this.state.errorMessage}
                      </Label>
                      <DefaultButton
                        text="Back"
                        title="Back"
                        iconProps={backIcon}
                        onClick={() => this.props.onClickCancel()}
                        styles={backBtnStyles}>
                      </DefaultButton>
                  </div>
                  )
                  :
                  <Label className={styles.errorMessage}>
                    {this.state.errorMessage}
                  </Label>
                }
                </div>
              )}
            </div>

            {this.state.tournamentName != "" && (
              <div>
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
                        <span className={styles.labelHeading}>Description</span>:
                        <span className={styles.labelNormal}>
                          {this.state.tournamentDescription}
                        </span>
                      </li>
                    )}
                  </ul>
                )}
              </div>
            )}
            {this.state.tournamentName != "" && (
              <div className={styles.contentArea}>
                <Row>
                  <Col>
                    <Label className={styles.subHeaderUnderline}>
                      Pending Actions
                    </Label>
                    {this.state.noPendingActions && (
                      <Label className={styles.successMessage}>
                        <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
                        There are no more pending actions in this tournament.
                      </Label>
                    )}
                    <TreeView
                      items={this.state.actionsList}
                      showCheckboxes={true}
                      selectChildrenIfParentSelected={true}
                      selectionMode={TreeViewSelectionMode.Multiple}
                      defaultExpanded={true}
                      onSelect={this.onActionSelected}
                    />
                    {this.state.actionsError && (
                      <Label className={styles.errorMessage}>
                        Select atleast one action to proceed.
                      </Label>
                    )}
                    {this.state.showSpinner && (
                      <Spinner
                        label={stringsConstants.formSavingMessage}
                        size={SpinnerSize.large}
                      />
                    )}
                    <div className={styles.btnArea}>
                      {this.state.actionsList.length != 0 && (
                        <PrimaryButton
                          text="Save"
                          title="Save"
                          iconProps={saveIcon}
                          onClick={this.saveActions}
                          className={styles.saveBtn}
                        ></PrimaryButton>
                      )}
                      &nbsp; &nbsp;
                      <PrimaryButton
                        text="Back"
                        title="Back"
                        iconProps={backIcon}
                        onClick={() => this.props.onClickCancel()}
                        styles={backBtnStyles}
                      ></PrimaryButton>
                    </div>
                  </Col>
                  <Col>
                    <Label className={styles.subHeaderUnderline}>
                      Completed Actions
                    </Label>
                    <TreeView
                      items={this.state.completedActionsList}
                      defaultExpanded={true}
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
