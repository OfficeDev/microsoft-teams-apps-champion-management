import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTEnableTournament.module.scss";

//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";

//Fluent UI controls
import { IButtonStyles, IChoiceGroupStyles, PrimaryButton, Spinner, SpinnerSize } from "@fluentui/react";
import { Label } from "@fluentui/react/lib/Label";
import { ChoiceGroup, IChoiceGroupOption } from "@fluentui/react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { IIconProps } from '@fluentui/react/lib/Icon';
import * as LocaleStrings from 'ClbHomeWebPartStrings';

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
  activeTournamentsList: any;
  selectedTournament: string;
  selectedTournamentId: string;
  selectedActiveTournament: string;
  selectedActiveTournamentId: string;
  activeTournamentId: string;
  activeTournamentFlag: boolean;
  showSuccess: boolean;
  successMessage: string;
  showError: boolean;
  errorMessage: string;
  startTournamentError: boolean;
  endTournamentError: boolean;
  hideDialog: boolean;
  noTournamentsFlag: boolean;
  tournamentAction: string;
  showSpinner: boolean;
}

const endButtonStyles: Partial<IButtonStyles> = {
  root: {
    marginTop: "32px",
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
          color: "#727170",
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
      activeTournamentsList: [],
      selectedTournament: "",
      selectedTournamentId: "",
      selectedActiveTournament: "",
      selectedActiveTournamentId: "",
      activeTournamentId: "",
      activeTournamentFlag: true,
      showSuccess: false,
      showError: false,
      errorMessage: "",
      startTournamentError: false,
      endTournamentError: false,
      successMessage: "",
      hideDialog: true,
      noTournamentsFlag: false,
      tournamentAction: "",
      showSpinner: false,

    };
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
    //Bind Methods
    this.getPendingTournaments = this.getPendingTournaments.bind(this);
    this.onTournamentSelect = this.onTournamentSelect.bind(this);
    this.onActiveTournamentSelect = this.onActiveTournamentSelect.bind(this);
    this.enableTournament = this.enableTournament.bind(this);
    this.getActiveTournaments = this.getActiveTournaments.bind(this);
    this.completeTournament = this.completeTournament.bind(this);
    this.startTournament = this.startTournament.bind(this);
    this.endTournament = this.endTournament.bind(this);
  }

  //On load of app bind active and pending tournaments to choice lists
  public async componentDidMount() {
    this.getPendingTournaments();
    this.getActiveTournaments();
  }

  //get active tournament from Tournaments list and populate the label
  private async getActiveTournaments() {
    console.log(
      stringsConstants.TotLog + "Getting active tournament from Tournaments list."
    );
    try {
      let filterActive: string =
        "Status eq '" + stringsConstants.TournamentStatusActive + "'";
      const activeTournamentsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentsMasterList,
          filterActive
        );

      var activeTournamentsChoices = [];
      if (activeTournamentsArray.length > 0) {
        //Loop through all "Active" tournaments and create an array with key and text
        await activeTournamentsArray.forEach((eachTournament) => {
          activeTournamentsChoices.push({
            key: eachTournament["Id"],
            text: eachTournament["Title"],
          });
        });

        this.setState({ activeTournamentsList: activeTournamentsChoices });
      }
      //If no active tournaments are found in the Tournaments list, set the flag
      else
        this.setState({ activeTournamentFlag: false });

    } catch (error) {
      console.error("TOT_TOTEnableTournament_getActiveTournaments \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while retrieving the active tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  // Get a list of all tournaments that are in "Not Started" status and to bind to Choice List
  private async getPendingTournaments() {
    console.log(
      stringsConstants.TotLog + "Getting tournaments with 'Not Started' status from Tournaments list."
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
      //If no pending tournaments are found in the Tournaments list, set the flag
      else this.setState({ noTournamentsFlag: true });
    } catch (error) {
      console.error("TOT_TOTEnableTournament_getPendingTournaments \n", error);
      this.setState({
        showError: true,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while retrieving the tournaments list. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //on select of a tournament, set the state
  private onTournamentSelect = async (
    ev: React.FormEvent<HTMLInputElement>,
    option: IChoiceGroupOption
  ): Promise<void> => {
    this.setState({
      selectedTournament: option.text,
      selectedTournamentId: option.key,
    });
  }

  //on select of an active tournament, set the state
  private onActiveTournamentSelect = async (
    ev: React.FormEvent<HTMLInputElement>,
    option: IChoiceGroupOption
  ): Promise<void> => {
    this.setState({
      selectedActiveTournament: option.text,
      selectedActiveTournamentId: option.key,
    });
  }

  //Start Tournament

  //Show popup for starting a tournament
  private startTournament() {
    //clear previous error messages on the form
    this.setState({
      showError: false,
      startTournamentError: false,
      hideDialog: true,
      errorMessage: "",
      successMessage: "",
      showSuccess: false,
    });

    if (this.state.selectedTournament == "")
      this.setState({ startTournamentError: true });
    else
      this.setState({ hideDialog: false, tournamentAction: stringsConstants.StartTournamentAction });
  }

  //On enabling a tournament change the status in Tournaments list
  private async enableTournament() {
    try {
      //clear previous error messages on the form
      this.setState({
        showError: false,
        startTournamentError: false,
        hideDialog: true,
      });
      console.log(stringsConstants.TotLog + "Enabling selected tournament.");
      if (this.state.selectedTournament == "")
        this.setState({ startTournamentError: true });
      else {
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
            //Set selected tournament as Active tournament once enabled

            let addSelectedTournament: any[] = this.state.activeTournamentsList;
            addSelectedTournament.push({
              key: this.state.selectedTournamentId,
              text: this.state.selectedTournament,
            });
            this.setState({ activeTournamentsList: addSelectedTournament, activeTournamentFlag: true, });

            //Refresh the tournaments list after enabling a tournament by deleting it from the array
            let newTournamentsRefresh: any[] = this.state.tournamentsList;

            var removeIndex = newTournamentsRefresh.map((item) => {
              return item.text;
            }).indexOf(this.state.selectedTournament);

            newTournamentsRefresh.splice(removeIndex, 1);
            this.setState({ tournamentsList: newTournamentsRefresh });

            if (newTournamentsRefresh.length == 0) {
              this.setState({ noTournamentsFlag: true });
            }

            //Show success message and clear state variable
            this.setState({
              showSuccess: true,
              successMessage: LocaleStrings.EnableTournamentSuccessMessage,
              selectedTournament: "",
              selectedTournamentId: ""
            });
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

  //Show popup for ending a tournament
  private endTournament() {
    //clear previous error messages on the form
    this.setState({
      showError: false,
      endTournamentError: false,
      hideDialog: true,
      showSuccess: false,
      successMessage: "",
      errorMessage: ""
    });
    if (this.state.selectedActiveTournament == "")
      this.setState({ endTournamentError: true });
    else
      this.setState({ hideDialog: false, tournamentAction: stringsConstants.EndTournamentAction });
  }

  //On Completing a tournament change the status in the Tournaments list
  private async completeTournament() {
    try {
      //clear previous error messages on the form
      this.setState({
        showError: false,
        endTournamentError: false,
        hideDialog: true,
        showSpinner: true,
      });
      console.log(stringsConstants.TotLog + "Completing active tournament.");

      if (this.state.selectedActiveTournament == "") {
        this.setState({
          endTournamentError: true,
          showSpinner: false
        });
      }
      else {

        let submitTournamentObject: any = {
          Status: stringsConstants.TournamentStatusCompleted,
        };
        //Creating items in Participants Report and Tournaments Report List for the completed tournament
        await commonServiceManager
          .updateCompletedTournamentDetails(
            this.state.selectedActiveTournament, new Date()
          ).then(() => {
            //Updating status for the completed tournament in Tournaments List
            commonServiceManager
              .updateListItem(
                stringsConstants.TournamentsMasterList,
                submitTournamentObject,
                this.state.selectedActiveTournamentId
              ).then(() => {
                //Refresh the active tournaments list after completing a tournament by deleting it from the array
                let newActiveTournamentsRefresh: any[] = this.state.activeTournamentsList;

                var removeIndex = newActiveTournamentsRefresh.map((item) => {
                  return item.text;
                }).indexOf(this.state.selectedActiveTournament);

                newActiveTournamentsRefresh.splice(removeIndex, 1);
                this.setState({ activeTournamentsList: newActiveTournamentsRefresh });

                if (newActiveTournamentsRefresh.length == 0)
                  this.setState({ activeTournamentFlag: false, showSpinner: false });

                //Show success message and clear state for selected item
                this.setState({
                  selectedActiveTournament: "",
                  selectedActiveTournamentId: "",
                  showSuccess: true,
                  successMessage: LocaleStrings.EndTournamentSuccessMessage,
                  showSpinner: false
                });
              }).catch((error) => {
                console.error("TOT_TOTEnableTournament_completeTournament \n", error);
                this.setState({
                  showError: true,
                  showSpinner: false,
                  errorMessage:
                    stringsConstants.TOTErrorMessage +
                    "while completing the tournament. Below are the details: \n" +
                    JSON.stringify(error),
                });
              });
          }).catch((error) => {
            console.error("TOT_TOTEnableTournament_completeTournament \n", error);
            this.setState({
              showError: true,
              errorMessage:
                stringsConstants.TOTErrorMessage +
                "while completing the tournament. Below are the details: \n" +
                JSON.stringify(error),
            });
          });
      }
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
          <span className={styles.manageTournamentLabel}>{LocaleStrings.ManageTournamentsPageTitle}</span>
        </div>
        <h5 className={styles.pageHeader}>{LocaleStrings.ManageTournamentsPageTitle}</h5>
        <div className={styles.textLabels}>
          <Label>{LocaleStrings.ManageToTLabel1}</Label>
          <Label>{LocaleStrings.ManageToTLabel2}</Label>
        </div>
        <br />
        {!this.state.hideDialog && (
          <Dialog
            hidden={this.state.hideDialog}
            onDismiss={() => this.setState({ hideDialog: true })}
            dialogContentProps={{
              type: DialogType.close,
              title: LocaleStrings.ConfirmLabel,
              subText: this.state.tournamentAction == stringsConstants.StartTournamentAction ? LocaleStrings.StartTournamentDialogMessage : LocaleStrings.EndTournamentDialogMessage,
              className: styles.endTotDialougeText
            }}
            containerClassName={'ms-dialogMainOverride ' + styles.textDialog}
            modalProps={{ isBlocking: false }}
          >
            <DialogFooter>
              {this.state.tournamentAction == stringsConstants.EndTournamentAction && (
                <PrimaryButton
                  onClick={this.completeTournament}
                  text={LocaleStrings.YesButton}
                  className={styles.yesBtn}
                  title={LocaleStrings.YesButton}
                />
              )}
              {this.state.tournamentAction == stringsConstants.StartTournamentAction && (
                <PrimaryButton
                  onClick={this.enableTournament}
                  text={LocaleStrings.YesButton}
                  className={styles.yesBtn}
                  title={LocaleStrings.YesButton}
                />
              )}
              <PrimaryButton
                onClick={() => this.setState({ hideDialog: true })}
                text={LocaleStrings.NoButton}
                className={styles.noBtn}
                title={LocaleStrings.NoButton}
              />
            </DialogFooter>
          </Dialog>
        )}
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
          <div className={styles.tournamentStatus}>
            <h4 className={styles.subHeaderUnderline}>{LocaleStrings.ActiveTournamentLabel}</h4>
          </div>
          <Row>
            <Col md={6}>
              {this.state.activeTournamentsList.length !== 0 &&
                <ChoiceGroup
                  onChange={this.onActiveTournamentSelect.bind(this)}
                  options={this.state.activeTournamentsList}
                  styles={choiceGroupStyles}
                />
              }
              {!this.state.activeTournamentFlag && (
                <Label className={styles.errorMessage}>
                  {LocaleStrings.NoActiveTournamentMessage}
                </Label>
              )}
              {this.state.endTournamentError && (
                <Label className={styles.errorMessage}>
                  {LocaleStrings.SelectEndTournamentMessage}
                </Label>
              )}
              {this.state.showSpinner &&
                <Spinner
                  label={LocaleStrings.CompleteTournamentSpinnerMessage}
                  size={SpinnerSize.large}
                />
              }
            </Col>
          </Row>
          <Row>
            <Col md={6}>
              {this.state.activeTournamentFlag && (
                <PrimaryButton
                  iconProps={{
                    iconName: "StatusCircleErrorX"
                  }}
                  text={LocaleStrings.EndTournamentButton}
                  title={LocaleStrings.EndTournamentButton}
                  styles={endButtonStyles}
                  onClick={this.endTournament}
                  className={this.state.showSpinner ? styles.disabledBtn : styles.completeBtn}
                  disabled={this.state.showSpinner}
                />
              )}
            </Col>
          </Row>
        </div>
        <br />
        <div className={styles.startTournmentArea}>
          <h4 className={styles.subHeaderUnderline}>{LocaleStrings.StartTournamentHeaderLabel}</h4>
        </div>
        <div>
          <Row>
            <Col md={6}>
              {this.state.tournamentsList !== 0 &&
                <ChoiceGroup
                  onChange={this.onTournamentSelect.bind(this)}
                  options={this.state.tournamentsList}
                  styles={choiceGroupStyles}
                />
              }
              {this.state.noTournamentsFlag && (
                <Label className={styles.errorMessage}>
                  {LocaleStrings.NoTournamentMessage}
                </Label>
              )}
              {this.state.startTournamentError && (
                <Label className={styles.errorMessage}>
                  {LocaleStrings.SelectTournamentMessage}
                </Label>
              )}
            </Col>
          </Row>
          <Row>
            <Col md={6}>
              <div className={styles.bottomBtnArea}>
                {!this.state.noTournamentsFlag && (
                  <PrimaryButton
                    text={LocaleStrings.StartTournamentButton}
                    title={LocaleStrings.StartTournamentButton}
                    iconProps={{
                      iconName: "Play"
                    }}
                    styles={startButtonStyles}
                    onClick={this.startTournament}
                    className={styles.enableBtn}
                  />
                )}
                &nbsp; &nbsp;
                <PrimaryButton
                  text={LocaleStrings.BackButton}
                  title={LocaleStrings.BackButton}
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
