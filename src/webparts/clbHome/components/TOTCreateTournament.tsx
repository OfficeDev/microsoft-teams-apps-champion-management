import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTCreateTournament.module.scss";

//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";

//FluentUI controls
import { TextField } from "@fluentui/react/lib/TextField";
import { IButtonStyles, PrimaryButton } from "@fluentui/react";
import { Label } from "@fluentui/react/lib/Label";
import { Icon, IIconProps } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';

//PNP
import {
  TreeView,
  ITreeItem,
  TreeViewSelectionMode,
} from "@pnp/spfx-controls-react/lib/TreeView";
import { ITextFieldStyles } from "office-ui-fabric-react/lib/components/TextField/TextField.types";

export interface ICreateTournamentProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
}

interface ICreateTournamentState {
  actionsList: ITreeItem[];
  tournamentName: string;
  tournamentDescription: string;
  selectedActionsList: ITreeItem[];
  tournamentError: Boolean;
  actionsError: Boolean;
  showForm: Boolean;
  showSuccess: Boolean;
  showError: Boolean;
  errorMessage: string;
}

const calloutProps = { gapSpace: 0 };

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', cursor: 'pointer' } };

const classes = mergeStyleSets({
  icon: {
    fontSize: '16px',
    paddingLeft: '10px',
    paddingTop: '6px',
    fontWeight: 'bolder',
    color: '#1d0f62'
  }
});

const labelStyles: Partial<ITextFieldStyles> = {
  subComponentStyles: {
    label: {
      root: {
        textAlign: "left",
        font: "normal normal 600 18px/24px Segoe UI",
        letterSpacing: "0px",
        color: "#000000",
        opacity: 1
      }
    }
  }
};

const backBtnStyles: Partial<IButtonStyles> = {
  root: {
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

const addIcon: IIconProps = { iconName: 'Add' };
const backIcon: IIconProps = { iconName: 'NavigateBack' };

export default class TOTCreateTournament extends React.Component<
  ICreateTournamentProps,
  ICreateTournamentState
> {
  constructor(props: ICreateTournamentProps, state: ICreateTournamentState) {
    super(props);
    //Set default values for state
    this.state = {
      actionsList: [],
      tournamentName: "",
      tournamentDescription: "",
      selectedActionsList: [],
      tournamentError: false,
      actionsError: false,
      showForm: true,
      showSuccess: false,
      showError: false,
      errorMessage: "",
    };

    //Bind Methods
    this.getActions = this.getActions.bind(this);
    this.handleInput = this.handleInput.bind(this);
    this.onActionSelected = this.onActionSelected.bind(this);
    this.saveTournament = this.saveTournament.bind(this);
  }

  //Get Actions from Master list and bind it to treeview on app load
  public componentDidMount() {
    //Get Actions from Master list and bind it to Treeview
    this.getActions();
  }

  //Get Actions from Master list and bind it to Treeview
  private async getActions() {
    console.log(stringsConstants.TotLog + "Getting actions from master list.");
    try {
      //Get all actions from 'Actions List'  to bind it to Treeview
      let commonServiceManager: commonServices = new commonServices(
        this.props.context,
        this.props.siteUrl
      );
      const allActionsArray: any[] = await commonServiceManager.getAllListItems(
        stringsConstants.ActionsMasterList
      );
      var treeItemsArray: ITreeItem[] = [];

      //Loop through all actions and build parent nodes(Categories) for Treeview
      await allActionsArray.forEach((vAction) => {
        const tree: ITreeItem = {
          key: vAction["Category"],
          label: vAction["Category"],
          children: [],
        };
        //Check if Category is already added to the Treeview. If yes, skip adding.
        var found = treeItemsArray.some((value) => {
          return value.label === vAction["Category"];
        });

        //Add category to Treeview only if it doesnt exists already.
        if (!found) treeItemsArray.push(tree);
      });

      //Loop through all actions and build child nodes(Actions) to the Treeview
      await allActionsArray.forEach((vAction) => {
        const tree: ITreeItem = {
          key: vAction.Id,
          label: vAction["Title"],
          data:
            vAction["Category"] +
            stringsConstants.StringSeperator +
            vAction["HelpURL"],
          subLabel:
            vAction["Points"] +
            stringsConstants.PointsDisplayString +
            vAction["Description"],
        };
        var treeCol: Array<ITreeItem> = treeItemsArray.filter((value) => {
          return value.label == vAction["Category"];
        });
        if (treeCol.length != 0) {
          treeCol[0].children.push(tree);
        }
      });
      this.setState({ actionsList: treeItemsArray });
    } catch (error) {
      console.error("TOT_TOTCreateTournament_getActions \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          " while retrieving actions list. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //Handle state values for form fields
  private handleInput(event: any, key: string) {
    switch (key) {
      case "tournamentName":
        this.setState({ tournamentName: event.target.value });
        break;
      case "tournamentDescription":
        this.setState({ tournamentDescription: event.target.value });
        break;
      default:
        break;
    }
  }

  //On select of a tree node change the state of selected actions
  private onActionSelected(items: ITreeItem[]) {
    this.setState({ selectedActionsList: items });
  }

  //Validate fields on the form and set a flag
  private ValidateFields(): Boolean {
    let validateFlag: Boolean = true;
    try {
      //clear previous error messages on the form
      this.setState({ showError: false });
      if (this.state.tournamentName == "") {
        validateFlag = false;
        this.setState({ tournamentError: true });
      } else this.setState({ tournamentError: false });
      if (this.state.selectedActionsList.length == 0) {
        validateFlag = false;
        this.setState({ actionsError: true });
      } else this.setState({ actionsError: false });
    } catch (error) {
      console.error("TOT_TOTCreateTournament_validateFields \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while validating the form. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
    return validateFlag;
  }

  // Save Tournament Details to SP Lists 'Tournaments' and 'Tournament Actions'
  private async saveTournament() {
    try {
      console.log(stringsConstants.TotLog + "saving tournament details.");
      let filter: string = "Title eq '" + this.state.tournamentName.trim().replace(/'/g, "''") + "'";
      if (this.ValidateFields()) {
        let commonServiceManager: commonServices = new commonServices(
          this.props.context,
          this.props.siteUrl
        );
        const allItems: any[] =
          await commonServiceManager.getItemsWithOnlyFilter(
            stringsConstants.TournamentsMasterList,
            filter
          );
        if (allItems.length == 0) {
          let submitTournamentsObject: any = {
            Title: this.state.tournamentName.trim(),
            Description: this.state.tournamentDescription,
            Status: stringsConstants.TournamentStatusNotStarted,
          };

          //Create item in 'Tournaments' list
          await commonServiceManager
            .createListItem(
              stringsConstants.TournamentsMasterList,
              submitTournamentsObject
            )
            .then((response) => {
              var selectedTreeArray: ITreeItem[] =
                this.state.selectedActionsList;
              //Loop through actions selected and create a list item for each treeview selection
              selectedTreeArray.forEach((c) => {
                //Skip parent node for treeview which is not an action
                if (c.data != undefined) {
                  let submitObject: any = {
                    Title: this.state.tournamentName.trim(),
                    Action: c.label,
                    Category: c.data.split(stringsConstants.StringSeperator)[0],
                    HelpURL: c.data.split(stringsConstants.StringSeperator)[1],
                    Points: c.subLabel
                      .split(stringsConstants.StringSeperatorPoints)[0]
                      .replace(stringsConstants.PointsReplaceString, ""),
                    Description: c.subLabel
                      .split(stringsConstants.StringSeperatorPoints)[1]
                      .replace(stringsConstants.PointsReplaceString, ""),
                  };

                  //Create an item in 'Tournament Actions' list for each selected action in tree view
                  commonServiceManager
                    .createListItem(
                      stringsConstants.TournamentActionsMasterList,
                      submitObject
                    )
                    .then((responseObj) => { })
                    .catch((error) => {
                      //Log error to console and display on the form
                      console.error(
                        "TOT_TOTCreateTournament_saveTournament \n",
                        JSON.stringify(error)
                      );
                      this.setState({
                        showError: true,
                        errorMessage:
                          stringsConstants.TOTErrorMessage +
                          "while saving the tournament action details to list. Below are the details: \n" +
                          JSON.stringify(error),
                        showForm: true,
                        showSuccess: false,
                      });
                    });
                }
              });
              this.setState({ showForm: false, showSuccess: true });
            })
            .catch((error) => {
              //Log error to console and display on the form
              console.error(
                "TOT_TOTCreateTournament_saveTournament \n",
                JSON.stringify(error)
              );
              this.setState({
                showError: true,
                errorMessage:
                  stringsConstants.TOTErrorMessage +
                  "while saving the tournament details to list. Below are the details: \n" +
                  JSON.stringify(error),
              });
            });
        } else {
          this.setState({
            showError: true,
            errorMessage:
              "Tournament name already exists. Enter another name for tournament.",
          });
        }
      }
    } catch (error) {
      console.error("TOT_TOTCreateTournament_saveTournament \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while saving the tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //Render Method
  public render(): React.ReactElement<ICreateTournamentProps> {
    return (
      <div className={styles.container}>
        <div className={styles.createTournamentPath}>
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
          <span className={styles.createTournamentLabel}>Create Tournament</span>
        </div>
        <div>
          {this.state.showSuccess && (
            <Label className={styles.successMessage}>
              <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
              Tournament created successfully.
            </Label>
          )}

          {this.state.showError && (
            <Label className={styles.errorMessage}>
              {this.state.errorMessage}
            </Label>
          )}
        </div>

        {this.state.showForm && (
          <div>
            <Row>
              <Col md={6}>
                <TextField
                  label="Tournament Name"
                  required
                  placeholder="Tournament Name"
                  maxLength={255}
                  value={this.state.tournamentName}
                  onChange={(evt) => this.handleInput(evt, "tournamentName")}
                  styles={labelStyles}
                />
                {this.state.tournamentError && (
                  <Label className={styles.errorMessage}>
                    Tournament Name is required.
                  </Label>
                )}
              </Col>
            </Row>
            <br />
            <Row>
              <Col md={6}>
                <TextField
                  label="Tournament Description"
                  multiline
                  maxLength={500}
                  placeholder="Tournament Description(Max 500 characters)"
                  value={this.state.tournamentDescription}
                  onChange={(evt) =>
                    this.handleInput(evt, "tournamentDescription")
                  }
                  styles={labelStyles}
                />
              </Col>
            </Row>
            <br />
            <Row>
              <Col className={styles.treeViewContent}>
                <div className={styles.selectTeamActionArea}>
                  <Label className={styles.selectTeamActionLabel}>
                    Select Teams Actions:{" "}
                    <span className={styles.asteriskStyle}>*</span>
                  </Label>
                  <TooltipHost
                    content="Select from the below available Teams actions to include in the new tournament. To add new tournament actions to choose from, visit the Manage Tournament Actions from the Admin Tools."
                    calloutProps={calloutProps}
                    styles={hostStyles}
                  >
                    <Icon aria-label="Info" iconName="Info" className={classes.icon} />

                  </TooltipHost>
                </div>
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
                    Select atleast one action to create a tournament.
                  </Label>
                )}
              </Col>
            </Row>
          </div>
        )}
        <div>
          <Row>
            <Col>
              {this.state.showForm && (
                <PrimaryButton
                  text="Create Tournament"
                  title="Create Tournament"
                  iconProps={addIcon}
                  onClick={this.saveTournament}
                  className={styles.createBtn}
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
            </Col>
          </Row>
        </div>
      </div> //Final DIV
    );
  }
}
