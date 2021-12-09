import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTEnableTournament.module.scss";
import * as strings from "../constants/strings";

//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";

//Fluent UI controls
import { IButtonStyles, IChoiceGroupStyles, PrimaryButton } from "@fluentui/react";
import { Label } from "@fluentui/react/lib/Label";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { Icon, IIconProps } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import { DirectionalHint } from "@microsoft/office-ui-fabric-react-bundle";

//Global Variables
let commonServiceManager: commonServices;
const backIcon: IIconProps = { iconName: 'NavigateBack' };

const backBtnStyles: Partial<IButtonStyles> = {
  root: {
    borderColor: "#33344A",
    backgroundColor: "white",
    height: "auto"
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

export interface IEnableTournamentProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
}

interface IEnableTournamentState {
  tournamentsList: any;
  selectedTournament: string;
  selectedTournamentId: string;
  activeTournament: string;
  activeTournamentId: string;
  activeTournamentFlag: boolean;
  showSuccess: boolean;
  successMessage: string;
  showError: boolean;
  errorMessage: string;
  tournamentError: boolean;
  hideDialog: boolean;
  noTournamentsFlag: boolean;
}

const tooltipStyles = {
  calloutProps: { gapSpace: 0, style: { paddingLeft: "4%" } }
};

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };

const classes = mergeStyleSets({
  icon: {
    fontSize: '16px',
    color: '#1d0f62',
    cursor: 'pointer',
    fontWeight: 'bolder',
  }
});

const endButtonStyles: Partial<IButtonStyles> = {
  root: {
    marginBottom: "20px",
    height: "38px"
  },
  textContainer: { fontSize: "16px" },
  icon: {
    fontSize: "26px",
    fontWeight: "bolder",
    color: "#FFFFFF",
    opacity: 1
  }
};

const startButtonStyles: Partial<IButtonStyles> = {
  root: {
    height: "38px"
  },
  textContainer: { fontSize: "16px" },
  icon: {
    fontSize: "18px",
    fontWeight: "bolder",
    color: "#FFFFFF",
    opacity: 1
  }
};

const choiceGroupStyles: Partial<IChoiceGroupStyles> = {
  flexContainer: [
    {
      selectors: {
        ".ms-ChoiceField": {
          textAlign: "left",
          font: "normal normal 600 18px/20px Segoe UI",
          letterSpacing: "0px",
          color: "#979593",
          opacity: 1
        }
      }
    }
  ]
};


export default class TOTEnableTournament extends React.Component<
  IEnableTournamentProps,
  IEnableTournamentState
> {
  constructor(props: IEnableTournamentProps, state: IEnableTournamentState) {
    super(props);
    //Set default values for state
    this.state = {
      tournamentsList: [],
      selectedTournament: "",
      activeTournament: "None",
      activeTournamentId: "",
      activeTournamentFlag: true,
      selectedTournamentId: "",
      showSuccess: false,
      showError: false,
      errorMessage: "",
      tournamentError: false,
      successMessage: "",
      hideDialog: true,
      noTournamentsFlag: false,
    };
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
    //Bind Methods
    this.getTournamentsList = this.getTournamentsList.bind(this);
    this.onTournamentSelect = this.onTournamentSelect.bind(this);
    this.enableTournament = this.enableTournament.bind(this);
    this.getActiveTournament = this.getActiveTournament.bind(this);
    this.completeTournament = this.completeTournament.bind(this);
  }

  //On load of app bind tournaments to choice list and populate the current Active tournament
  public async componentDidMount() {
    this.getTournamentsList();
    this.getActiveTournament();
  }

  //get active tournament from master list and populate the label
  private async getActiveTournament() {
    console.log(
      stringsConstants.TotLog + "Getting active tournament from master list."
    );
    try {
      let filterActive: string =
        "Status eq '" + stringsConstants.TournamentStatusActive + "'";
      const activeTournamentsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentsMasterList,
          filterActive
        );

      if (activeTournamentsArray.length > 0)
        this.setState({
          activeTournament: activeTournamentsArray[0]["Title"],
          activeTournamentId: activeTournamentsArray[0]["Id"],
        });
      else this.setState({ activeTournamentFlag: false });
    } catch (error) {
      console.error("TOT_TOTEnableTournament_getActiveTournament \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while retrieving the active tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  // Get a list of all tournaments that are "Not Started" and to bind to Choice List
  private async getTournamentsList() {
    console.log(
      stringsConstants.TotLog + "Getting tournaments from master list."
    );
    try {
      let selectFilter: string =
        "Status eq '" + stringsConstants.TournamentStatusNotStarted + "'";
      const allTournamentsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentsMasterList,
          selectFilter
        );
      var tournamentsChoices = [];
      if (allTournamentsArray.length > 0) {
        //Loop through all "Not Started" tournaments and create an array with key and text
        await allTournamentsArray.forEach((eachTournament) => {
          tournamentsChoices.push({
            key: eachTournament["Id"],
            text: eachTournament["Title"],
          });
        });

        this.setState({ tournamentsList: tournamentsChoices });
      }
      //If no tournaments are found in the master list set the flag
      else this.setState({ noTournamentsFlag: true });
    } catch (error) {
      console.error("TOT_TOTEnableTournament_getTournamentsList \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while retrieving the tournaments list. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //on select of a tournament set the state
  private onTournamentSelect = async (
    ev: React.FormEvent<HTMLInputElement>,
    option: IChoiceGroupOption
  ): Promise<void> => {
    this.setState({
      selectedTournament: option.text,
      selectedTournamentId: option.key,
    });
  }

  //On enabling a tournament change the status in master list
  private enableTournament() {
    try {
      //clear previous error messages on the form
      this.setState({
        showError: false,
        tournamentError: false,
        hideDialog: true,
      });
      if (this.state.selectedTournament == "")
        this.setState({ tournamentError: true });
      else {
        console.log(stringsConstants.TotLog + "Enabling selected tournament.");
        let submitTournamentObject: any = {
          Status: stringsConstants.TournamentStatusActive,
        };
        commonServiceManager
          .updateListItem(
            stringsConstants.TournamentsMasterList,
            submitTournamentObject,
            this.state.selectedTournamentId
          )
          .then((result) => {
            //Set Enabled tournament as Active tournament once enabled
            this.setState({
              activeTournamentFlag: true,
              activeTournament: this.state.selectedTournament,
              activeTournamentId: this.state.selectedTournamentId,
            });
            //clear the state values
            this.setState({ selectedTournament: "", selectedTournamentId: "" });
            //Show success message
            this.setState({
              showSuccess: true,
              successMessage: "Tournament enabled successfully.",
            });

            //Refresh the tournaments list after enabling a tournament by deleting it from the array
            let newTournamentsRefresh: any[] = this.state.tournamentsList;
            for (
              var counter = 0;
              counter < newTournamentsRefresh.length;
              counter++
            ) {
              if (
                newTournamentsRefresh[counter]["text"] ==
                this.state.activeTournament
              ) {
                newTournamentsRefresh.splice(counter, 1);
                this.setState({ tournamentsList: newTournamentsRefresh });
                break;
              }
            }
            if (newTournamentsRefresh.length == 0)
              this.setState({ noTournamentsFlag: true });
          })
          .catch((error) => {
            console.error("TOT_TOTEnableTournament_enableTournament \n", error);
            this.setState({
              showError: true,
              errorMessage:
                stringsConstants.TOTErrorMessage +
                "while enabling the tournament. Below are the details: \n" +
                JSON.stringify(error),
            });
          });
      }
    } catch (error) {
      console.error("TOT_TOTEnableTournament_enableTournament \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while enabling the tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //On Completing a tournament change the status in the master list
  private completeTournament() {
    try {
      //clear previous error messages on the form
      this.setState({
        showError: false,
        tournamentError: false,
        hideDialog: true,
      });
      console.log(stringsConstants.TotLog + "Completing active tournament.");

      let submitTournamentObject: any = {
        Status: stringsConstants.TournamentStatusCompleted,
      };

      commonServiceManager
        .updateListItem(
          stringsConstants.TournamentsMasterList,
          submitTournamentObject,
          this.state.activeTournamentId
        )
        .then((result) => {
          //Reset the state
          this.setState({
            activeTournamentFlag: false,
            activeTournament: "None",
            activeTournamentId: "",
            selectedTournament: "",
            selectedTournamentId: "",
          });
          //Show success message
          this.setState({
            showSuccess: true,
            successMessage: "Tournament ended successfully.",
          });
        })
        .catch((error) => {
          console.error("TOT_TOTEnableTournament_completeTournament \n", error);
          this.setState({
            showError: true,
            errorMessage:
              stringsConstants.TOTErrorMessage +
              "while completing the tournament. Below are the details: \n" +
              JSON.stringify(error),
          });
        });
    } catch (error) {
      console.error("TOT_TOTEnableTournament_completeTournament \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while completing the tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //Render Method
  public render(): React.ReactElement<IEnableTournamentProps> {
    return (
      <div className={`container ${styles.mainContainer}`}>
        <div className={styles.manageTournamentPath}>
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
          <span className={styles.manageTournamentLabel}>Manage Tournaments</span>
        </div>
        <h5 className={styles.pageHeader}>Manage Tournaments</h5>
        <div className={styles.textLabels}>
          <Label>{strings.ManageTotLabel1}</Label>
          <Label>{strings.ManageTotLabel2}</Label>
        </div>
        <br />
        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={() => this.setState({ hideDialog: true })}
          dialogContentProps={{
            type: DialogType.close,
            title: "Confirm",
            subText: this.state.activeTournamentFlag ? "Are you sure want to end the tournament?" : "Are you sure want to start the new tournament?",

          }}
          containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
          modalProps={{ isBlocking: false }}
        >
          <DialogFooter>
            {this.state.activeTournamentFlag && (
              <PrimaryButton
                onClick={this.completeTournament}
                text="Yes"
                className={styles.yesBtn}
                title="Yes"
              />
            )}
            {!this.state.activeTournamentFlag && (
              <PrimaryButton
                onClick={this.enableTournament}
                text="Yes"
                className={styles.yesBtn}
                title="Yes"
              />
            )}
            <PrimaryButton
              onClick={() => this.setState({ hideDialog: true })}
              text="No"
              className={styles.noBtn}
              title="No"
            />
          </DialogFooter>
        </Dialog>
        <div>
          {this.state.showSuccess && (
            <Label className={styles.successMessage}>
              <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
              {this.state.successMessage}
            </Label>
          )}

          {this.state.showError && (
            <Label className={styles.errorMessage}>
              {this.state.errorMessage}
            </Label>
          )}
        </div>
        <div>
          <Row>
            <Col>
              <div className={styles.tournamentStatus}>
                <label>Active Tournament: </label>&nbsp;
                <label> {this.state.activeTournament} </label>
              </div>
            </Col>
          </Row>
          <Row>
            <Col>
              <PrimaryButton
                disabled={!this.state.activeTournamentFlag}
                iconProps={{
                  iconName: "StatusCircleErrorX"
                }}
                text="End Tournament"
                title="End Tournament"
                styles={endButtonStyles}
                onClick={() => this.setState({ hideDialog: false })}
                className={
                  !this.state.activeTournamentFlag
                    ? styles.disabledBtn
                    : styles.completeBtn
                }
              />
            </Col>
          </Row>
        </div>
        <br />
        <div className={styles.startTournmentArea}>
          <h4 className={styles.subHeaderUnderline}>Start Tournament{" "}</h4>
          <div className={styles.infoArea}>
            <TooltipHost
              content="A new tournament can only be started if there is no active tournament. To start a new tournament end the current tournament."
              tooltipProps={tooltipStyles}
              styles={hostStyles}
              directionalHint={DirectionalHint.bottomCenter}
            >
              <Icon aria-label="Info" iconName="Info" className={classes.icon} />
            </TooltipHost>
          </div>
        </div>
        <div>
          <Row>
            <Col md={6}>
              <ChoiceGroup
                disabled={this.state.activeTournamentFlag}
                onChange={this.onTournamentSelect.bind(this)}
                options={this.state.tournamentsList}
                styles={choiceGroupStyles}

              />
              {this.state.noTournamentsFlag && (
                <Label className={styles.errorMessage}>
                  No tournaments found.
                </Label>
              )}
              {this.state.tournamentError && (
                <Label className={styles.errorMessage}>
                  Select a tournament to enable.
                </Label>
              )}
            </Col>
          </Row>
          <Row>
            <Col md={6}>
              <div className={styles.bottomBtnArea}>
                {!this.state.noTournamentsFlag && (
                  <PrimaryButton
                    disabled={this.state.activeTournamentFlag}
                    text="Start Tournament"
                    title="Start Tournament"
                    iconProps={{
                      iconName: "Play"
                    }}
                    styles={startButtonStyles}
                    onClick={() => this.setState({ hideDialog: false })}
                    className={
                      this.state.activeTournamentFlag
                        ? styles.disabledBtn
                        : styles.enableBtn
                    }
                  />
                )}
                &nbsp; &nbsp;
                <PrimaryButton
                  text="Back"
                  title="Back"
                  iconProps={backIcon}
                  styles={backBtnStyles}
                  onClick={() => this.props.onClickCancel()}
                />
              </div>
            </Col>
          </Row>
        </div>
      </div> //Final DIV
    );
  }
}
