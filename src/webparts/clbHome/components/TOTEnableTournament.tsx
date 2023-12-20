import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from "../Common/CommonServices";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/TOTEnableTournament.module.scss";

//React Boot Strap
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";

//Fluent UI controls
import { PrimaryButton, Spinner, SpinnerSize, TextField, ChoiceGroup, IChoiceGroupOption, Checkbox } from "@fluentui/react";
import { Label } from "@fluentui/react/lib/Label";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import * as LocaleStrings from 'ClbHomeWebPartStrings';

//Global Variables
let commonServiceManager: commonServices;
export interface IEnableTournamentProps {
  context?: WebPartContext;
  siteUrl: string;
  onClickCancel: Function;
  currentThemeName?: string;
}

interface IEnableTournamentState {
  tournamentsList: any;
  filteredPendingTournaments: any;
  selectedPendingTournaments: any;
  activeTournamentsList: any;
  filteredActiveTournaments: any;
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
  showActiveTournamentSpinner: boolean;
  showPendingTournamentSpinner: boolean;
  activeTournamentsSearchedText: string;
  pendingTournamentsSearchedText: string;
  showSelectAllLabel: boolean;
  selectAllChecked: boolean;
}
export default class TOTEnableTournament extends React.Component<IEnableTournamentProps, IEnableTournamentState> {

  constructor(props: IEnableTournamentProps) {
    super(props);

    //Set default values for state
    this.state = {
      tournamentsList: [],
      filteredPendingTournaments: [],
      selectedPendingTournaments: [],
      showSelectAllLabel: true,
      selectAllChecked: false,
      pendingTournamentsSearchedText: "",
      activeTournamentsList: [],
      filteredActiveTournaments: [],
      activeTournamentsSearchedText: "",
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
      showActiveTournamentSpinner: false,
      showPendingTournamentSpinner: false

    };
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
    //Bind Methods
    this.getPendingTournaments = this.getPendingTournaments.bind(this);
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
      this.setState({
        activeTournamentsList: [],
        filteredActiveTournaments: [],
        activeTournamentFlag: true
      });
      let filterActive: string =
        "Status eq '" + stringsConstants.TournamentStatusActive + "'";
      const activeTournamentsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentsMasterList,
          filterActive
        );

      let activeTournamentsChoices: any = [];
      if (activeTournamentsArray.length > 0) {
        //Loop through all "Active" tournaments and create an array with key and text
        activeTournamentsArray.forEach((eachTournament) => {
          activeTournamentsChoices.push({
            key: eachTournament["Id"],
            text: eachTournament["Title"],
            ariaLabel: eachTournament["Title"]
          });
        });
        activeTournamentsChoices.sort((a: any, b: any) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0));
        this.setState({
          activeTournamentsList: activeTournamentsChoices,
          filteredActiveTournaments: activeTournamentsChoices
        });
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
      this.setState({ tournamentsList: [], filteredPendingTournaments: [] });
      let selectFilter: string =
        "Status eq '" + stringsConstants.TournamentStatusNotStarted + "'";
      const allTournamentsArray: any[] =
        await commonServiceManager.getItemsWithOnlyFilter(
          stringsConstants.TournamentsMasterList,
          selectFilter
        );
      let tournamentsChoices: any = [];
      if (allTournamentsArray.length > 0) {
        //Loop through all "Not Started" tournaments and create an array with key and text
        allTournamentsArray.forEach((eachTournament) => {
          tournamentsChoices.push({
            key: eachTournament["Id"],
            text: eachTournament["Title"],
          });
        });
        tournamentsChoices.sort((a: any, b: any) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0));
        this.setState({ tournamentsList: tournamentsChoices, filteredPendingTournaments: tournamentsChoices });
      }
      //If no pending tournaments are found in the Tournaments list, set the flag
      else {
        this.setState({ noTournamentsFlag: true });
      }
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

    if (this.state.selectedPendingTournaments.length === 0)
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
        showPendingTournamentSpinner: true
      });
      console.log(stringsConstants.TotLog + "Enabling selected tournament(s).");
      if (this.state.selectedPendingTournaments.length === 0)
        this.setState({ startTournamentError: true, showPendingTournamentSpinner: false });
      else {
        let submitTournamentObject: any = {
          Status: stringsConstants.TournamentStatusActive,
        };
        let updateResponse = await commonServiceManager.updateMultipleItems(
          stringsConstants.TournamentsMasterList,
          submitTournamentObject,
          this.state.selectedPendingTournaments
        );

        if (updateResponse) {
          this.setState({
            showSuccess: true,
            successMessage: this.state.selectedPendingTournaments.length + " " + LocaleStrings.EnableTournamentSuccessMessage
          });
        }
        else {
          this.setState({
            showError: true,
            errorMessage: stringsConstants.TOTErrorMessage + "while enabling the tournament"
          });
        }
        //Show success message and clear state variable
        this.setState({
          showPendingTournamentSpinner: false,
          selectAllChecked: false,
          showSelectAllLabel: true,
          selectedPendingTournaments: [],
          pendingTournamentsSearchedText: ""
        });
        await this.getActiveTournaments();
        await this.getPendingTournaments();
      }

    } catch (error) {

      await this.getActiveTournaments();
      await this.getPendingTournaments();

      this.setState({
        showError: true,
        showPendingTournamentSpinner: false,
        selectAllChecked: false,
        showSelectAllLabel: true,
        selectedPendingTournaments: [],
        pendingTournamentsSearchedText: "",
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while enabling the tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
      console.error("TOT_TOTEnableTournament_enableTournament \n", JSON.stringify(error));
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
        showActiveTournamentSpinner: true,
      });
      console.log(stringsConstants.TotLog + "Completing active tournament.");

      if (this.state.selectedActiveTournament == "") {
        this.setState({
          endTournamentError: true,
          showActiveTournamentSpinner: false
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

                let removeIndex = newActiveTournamentsRefresh.map((item) => {
                  return item.text;
                }).indexOf(this.state.selectedActiveTournament);

                newActiveTournamentsRefresh.splice(removeIndex, 1);
                newActiveTournamentsRefresh.sort((a: any, b: any) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0));
                this.setState({ activeTournamentsList: newActiveTournamentsRefresh, filteredActiveTournaments: newActiveTournamentsRefresh });

                if (newActiveTournamentsRefresh.length == 0)
                  this.setState({ activeTournamentFlag: false, showActiveTournamentSpinner: false });

                //Show success message and clear state for selected item
                this.setState({
                  selectedActiveTournament: "",
                  selectedActiveTournamentId: "",
                  showSuccess: true,
                  activeTournamentsSearchedText: "",
                  successMessage: LocaleStrings.EndTournamentSuccessMessage,
                  showActiveTournamentSpinner: false
                });
              }).catch((error) => {
                console.error("TOT_TOTEnableTournament_completeTournament \n", error);
                this.setState({
                  showError: true,
                  showActiveTournamentSpinner: false,
                  activeTournamentsSearchedText: "",
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
              showActiveTournamentSpinner: false,
              activeTournamentsSearchedText: "",
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
        showActiveTournamentSpinner: false,
        errorMessage:
          stringsConstants.TOTErrorMessage +
          "while completing the tournament. Below are the details: \n" +
          JSON.stringify(error),
      });
    }
  }

  //Search Active Tournaments to complete
  public searchActiveTournaments = (eve: any, value: string) => {
    this.setState({ activeTournamentsSearchedText: value });
    const trimmedValue = value.trim();
    if (trimmedValue !== "") {
      const searchedTournaments = this.state.activeTournamentsList.filter((tournamentData: any) => {
        return tournamentData.text.trim().toLowerCase().indexOf(trimmedValue.toLowerCase()) !== -1;
      });
      searchedTournaments.sort((a: any, b: any) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0));
      this.setState({ filteredActiveTournaments: searchedTournaments });
    }
    else {
      this.setState({ filteredActiveTournaments: this.state.activeTournamentsList });
    }
  }

  //Search Pending tournaments to start
  public searchPendingTournaments = (eve: any, value: string) => {
    this.setState({ pendingTournamentsSearchedText: value });
    const trimmedValue = value.trim();
    if (trimmedValue !== "") {
      this.setState({ showSelectAllLabel: false });
      const searchedTournaments = this.state.tournamentsList.filter((tournamentData: any) => {
        return tournamentData.text.trim().toLowerCase().indexOf(trimmedValue.toLowerCase()) !== -1;
      });
      searchedTournaments.sort((a: any, b: any) => (a.text > b.text) ? 1 : ((b.text > a.text) ? -1 : 0));
      this.setState({ filteredPendingTournaments: searchedTournaments });
    }
    else {
      this.setState({ showSelectAllLabel: true, filteredPendingTournaments: this.state.tournamentsList });
    }
  }

  //Update all selected pending tournaments to new array
  public updateSelectedPendingTournaments = (isChecked: boolean, key: any, selectAll: boolean) => {
    //When "Select All" is checked
    if (selectAll && isChecked) {
      this.setState({ selectAllChecked: true });
      let selectedTournaments = [];
      for (let tournamentObj of this.state.tournamentsList) {
        selectedTournaments.push(tournamentObj.key);
      }
      this.setState({ selectedPendingTournaments: selectedTournaments });
    }
    // When "Select All" is unchecked
    else if (selectAll && !isChecked) {
      this.setState({ selectAllChecked: false, selectedPendingTournaments: [] });
    }
    else {
      //When checkbox is checked
      if (isChecked) {
        let selectedTournaments = this.state.selectedPendingTournaments;
        selectedTournaments.push(key);
        this.setState({ selectedPendingTournaments: selectedTournaments });
        //Automatically check the "Select All" option when the last checkbox is checked
        if (selectedTournaments.length === this.state.tournamentsList.length) {
          this.setState({ selectAllChecked: true });
        }

      }
      //When checkbox is unchecked
      else {
        const selectedTournaments = this.state.selectedPendingTournaments.filter((tKey: any) => {
          return tKey !== key;
        });
        this.setState({
          selectAllChecked: false,
          selectedPendingTournaments: selectedTournaments
        });
      }
    }
  }

  //Render Method
  public render(): React.ReactElement<IEnableTournamentProps> {
    const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
    return (
      <div className={`container ${styles.manageTournamentsWrapper}${isDarkOrContrastTheme ? " " + styles.manageTournamentsWrapperDarkContrast : ""}`}>
        <div className={styles.manageTournamentPath}>
          <img src={require("../assets/CMPImages/BackIcon.png")}
            className={styles.backImg}
            alt={LocaleStrings.BackButton}
            aria-hidden="true"
          />
          <span
            className={styles.backLabel}
            onClick={(!this.state.showActiveTournamentSpinner || !this.state.showPendingTournamentSpinner) && (() => this.props.onClickCancel())}
            role="button"
            tabIndex={0}
            onKeyDown={(!this.state.showActiveTournamentSpinner || !this.state.showPendingTournamentSpinner) && ((evt: any) => { if (evt.key === stringsConstants.stringEnter || evt.key === stringsConstants.stringSpace) this.props.onClickCancel() })}
            aria-label={LocaleStrings.TOTBreadcrumbLabel}
          >
            <span title={LocaleStrings.TOTBreadcrumbLabel}>
              {LocaleStrings.TOTBreadcrumbLabel}
            </span>
          </span>
          <span className={styles.border} aria-live="polite" role="alert" aria-label={LocaleStrings.ManageTournamentsPageTitle + " Page"} />
          <span className={styles.manageTournamentLabel}>{LocaleStrings.ManageTournamentsPageTitle}</span>
        </div>
        <h2 className={styles.pageHeader} role="heading" tabIndex={0}>{LocaleStrings.ManageTournamentsPageTitle}</h2>
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
            modalProps={{ isBlocking: false, className: styles.textDialog }}
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
            <Label className={styles.successMessage} aria-live="polite" role="alert">
              <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" aria-hidden="true" className={styles.tickImage} />
              {this.state.successMessage}
            </Label>
          )}
          {this.state.showError && (
            <Label className={styles.errorMessage} aria-live="polite" role="alert">
              {this.state.errorMessage}
            </Label>
          )}
        </div>
        <Row xl={2} lg={2} md={2} sm={1} xs={1}>
          <Col xl={5} lg={5} md={6} sm={12} xs={12}>
            <div className={styles.tournamentStatus}>
              <h3 className={styles.subHeaderUnderline} role="heading" tabIndex={0}>{LocaleStrings.ActiveTournamentLabel}</h3>
            </div>
            {this.state.activeTournamentsList.length !== 0 &&
              <>
                <TextField
                  className={styles.manageTrmntSearchBox}
                  placeholder={LocaleStrings.SearchActiveTournaments}
                  onChange={this.searchActiveTournaments}
                  iconProps={{
                    iconName: this.state.activeTournamentsSearchedText !== "" ? "ChromeClose" : "Search",
                    className: `${styles.clearSearchIcon} ${this.state.activeTournamentsSearchedText !== "" ? styles.chromeCloseIcon : styles.searchIcon}`,
                    onClick: this.state.activeTournamentsSearchedText !== "" ? () => {
                      this.setState({
                        activeTournamentsSearchedText: "",
                        filteredActiveTournaments: this.state.activeTournamentsList
                      });
                    } : null
                  }}
                  value={this.state.activeTournamentsSearchedText}
                  disabled={this.state.showActiveTournamentSpinner}
                />
                {this.state.filteredActiveTournaments.length > 0 && (navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
                  <span aria-live="polite" role="alert">
                    {this.state.filteredActiveTournaments.length}&nbsp;{stringsConstants.activeTournamentsFoundLabel}
                  </span>
                }
                <div className={styles.manageActiveTrmntChoiceGrpArea}>
                  <ChoiceGroup
                    onChange={this.onActiveTournamentSelect.bind(this)}
                    options={this.state.filteredActiveTournaments}
                    className={styles.manageActiveTrmntChoiceGrp}
                    id="activeTournamentsList"
                    disabled={this.state.showActiveTournamentSpinner}
                  />
                  {this.state.filteredActiveTournaments.length === 0 &&
                    <div className={styles.noResultsFound} aria-live="polite" role="alert">{LocaleStrings.NoSearchResults}</div>
                  }
                </div>
              </>
            }
            {this.state.filteredActiveTournaments.length > 0 && !(navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
              <span aria-live="polite" role="alert"
                aria-label={this.state.filteredActiveTournaments.length + " " + stringsConstants.activeTournamentsFoundLabel} />
            }
            {!this.state.activeTournamentFlag && (
              <Label className={styles.errorMessage} tabIndex={0} role="status">
                {LocaleStrings.NoActiveTournamentMessage}
              </Label>
            )}
            {this.state.endTournamentError && (
              <Label className={styles.errorMessage} id="end-tournament-error" role="status">
                {LocaleStrings.SelectEndTournamentMessage}
              </Label>
            )}
            {this.state.showActiveTournamentSpinner &&
              <Spinner
                label={LocaleStrings.CompleteTournamentSpinnerMessage}
                size={SpinnerSize.large}
              />
            }
            {this.state.activeTournamentFlag && (
              <PrimaryButton
                iconProps={{ iconName: "CalculatorMultiply" }}
                text={LocaleStrings.EndTournamentButton}
                title={LocaleStrings.EndTournamentButton}
                onClick={this.endTournament}
                className={styles.endTrmntBtn}
                disabled={this.state.showActiveTournamentSpinner}
                aria-describedby="end-tournament-error"
                tabIndex={0}
              />
            )}
          </Col>
          <Col xl={5} lg={5} md={6} sm={12} xs={12}>
            <div className={styles.tournamentStatus}>
              <h3 className={styles.subHeaderUnderline} role="heading" tabIndex={0}>{LocaleStrings.StartTournamentHeaderLabel}</h3>
            </div>
            <div>
              {this.state.tournamentsList.length !== 0 &&
                <div>
                  <TextField
                    className={styles.manageTrmntSearchBox}
                    placeholder={LocaleStrings.SearchPendingTournaments}
                    onChange={this.searchPendingTournaments}
                    iconProps={{
                      iconName: this.state.pendingTournamentsSearchedText !== "" ? "ChromeClose" : "Search",
                      className: `${styles.clearSearchIcon} ${this.state.pendingTournamentsSearchedText !== "" ? styles.chromeCloseIcon : styles.searchIcon}`,
                      onClick: this.state.pendingTournamentsSearchedText !== "" ? () => {
                        this.setState({
                          pendingTournamentsSearchedText: "",
                          showSelectAllLabel: true,
                          filteredPendingTournaments: this.state.tournamentsList
                        });
                      } : null
                    }}
                    value={this.state.pendingTournamentsSearchedText}
                    disabled={this.state.showPendingTournamentSpinner}
                  />
                  {this.state.filteredPendingTournaments.length > 0 && (navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
                    <span aria-live="polite" role="alert">
                      {this.state.filteredPendingTournaments.length}&nbsp;{stringsConstants.newTournamentsFoundLabel}
                    </span>
                  }
                  <div className={styles.managePendingTrmntCheckboxGrp}>
                    {this.state.showSelectAllLabel &&
                      <Checkbox
                        label={LocaleStrings.SelectAllLabel}
                        onChange={(_eve: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked: boolean) => {
                          this.updateSelectedPendingTournaments(isChecked, "", true);
                        }}
                        className={styles.pendingTrmntCheckBox}
                        checked={this.state.selectAllChecked}
                        disabled={this.state.showPendingTournamentSpinner}
                        ariaLabel={LocaleStrings.SelectAllLabel}
                        key="select-all-checkbox"
                      />
                    }
                    {this.state.filteredPendingTournaments.map((tournamentData: any) => {
                      return (
                        <Checkbox
                          label={tournamentData.text}
                          onChange={(_, isChecked: boolean) => {
                            this.updateSelectedPendingTournaments(isChecked, tournamentData.key, false);
                          }}
                          checked={this.state.selectedPendingTournaments.includes(tournamentData.key)}
                          className={styles.pendingTrmntCheckBox + " " + styles.filteredPendingTournaments}
                          disabled={this.state.showPendingTournamentSpinner}
                          ariaLabel={tournamentData.text}
                          key={tournamentData.key}
                        />
                      );
                    })}
                    {this.state.filteredPendingTournaments.length === 0 &&
                      <div className={styles.noResultsFound} role="alert" aria-live="polite">
                        {LocaleStrings.NoSearchResults}</div>
                    }
                  </div>
                </div>
              }
            </div>
            {this.state.filteredPendingTournaments.length > 0 && !(navigator.userAgent.match(/iPhone/i) || navigator.userAgent.match(/iPad/i)) &&
              <span aria-live="polite" role="alert"
                aria-label={this.state.filteredPendingTournaments.length + " " + stringsConstants.newTournamentsFoundLabel} />
            }
            {this.state.noTournamentsFlag && (
              <Label className={styles.errorMessage} tabIndex={0} role="status">
                {LocaleStrings.NoTournamentMessage}
              </Label>
            )}
            {this.state.startTournamentError && (
              <Label className={styles.errorMessage} id="start-tournament-error" role="status">
                {LocaleStrings.SelectTournamentMessage}
              </Label>
            )}
            {this.state.showPendingTournamentSpinner &&
              <Spinner
                label={LocaleStrings.EnableTournamentSpinnerMessage}
                size={SpinnerSize.large}
              />
            }
            {!this.state.noTournamentsFlag && (
              <PrimaryButton
                text={LocaleStrings.StartTournamentButton}
                title={LocaleStrings.StartTournamentButton}
                iconProps={{ iconName: "Play" }}
                onClick={(!this.state.showActiveTournamentSpinner || !this.state.showPendingTournamentSpinner) && this.startTournament}
                className={styles.enableBtn}
                disabled={this.state.showPendingTournamentSpinner}
                aria-describedby="start-tournament-error"
                tabIndex={0}
              />
            )}
          </Col>
        </Row>
        <PrimaryButton
          text={LocaleStrings.BackButton}
          title={LocaleStrings.BackButton}
          iconProps={{ iconName: 'NavigateBack' }}
          onClick={() => this.props.onClickCancel()}
          className={styles.manageTrmntBackBtn}
          tabIndex={0}
        />
      </div> //Final DIV
    );
  }
}
