import { Icon } from '@fluentui/react/lib/Icon';
import {
  ISPHttpClientOptions, SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import {
  PeoplePicker,
  PrincipalType
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as React from "react";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/CMPAddMember.module.scss";
import commonServices from "../Common/CommonServices";
import ClbChampionsList from './ClbChampionsList';
import { IConfigList } from './ManageConfigSettings';


//global variables
let commonServiceManager: commonServices;
export interface IClbAddMemberProps {
  context?: any;
  onClickBack: () => void;
  onHomeCallBack: () => void;
  siteUrl: string;
  isAdmin: boolean;
  appTitle: string;
  currentThemeName?: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  FirstName: string;
  LastName: string;
  Country: String;
  Status: String;
  Role: String;
  Region: string;
  Points: number;
}
interface IUserDetail {
  ID: any;
  LoginName: string;
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  errorMessage: string;
  updatedMessage: string;
  UserDetails: Array<IUserDetail>;
  selectedusers: Array<any>;
  siteUrl: string;
  countries: Array<any>;
  regions: Array<any>;
  groups: Array<any>;
  focusAreas: Array<any>;
  selectedFocusAreas: any;
  multiSelectChoices: any;
  memberData: any;
  memberrole: string;
  sitename: string;
  inclusionpath: string;
  load: boolean;
  memberListColumnNames: Array<any>;
  configListSettings: Array<IConfigList>;
  championsList: boolean;
  isUserAdded: boolean;
  userStatus: string;
  regionColumnName: string;
  countryColumnName: string;
  groupColumnName: string;
}
class ClbAddMember extends React.Component<IClbAddMemberProps, IState> {
  public addMemberPeoplePickerParentRef: React.RefObject<HTMLDivElement>;
  public addMemberPeoplePickerRef: React.RefObject<PeoplePicker>;
  constructor(props: IClbAddMemberProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });
    this.addMemberPeoplePickerParentRef = React.createRef();
    this.addMemberPeoplePickerRef = React.createRef();
    this.state = {
      list: { value: [] },
      isAddChampion: false,
      errorMessage: "",
      updatedMessage: "",
      UserDetails: [],
      selectedusers: [],
      countries: [],
      regions: [],
      groups: [],
      focusAreas: [],
      selectedFocusAreas: [],
      multiSelectChoices: [],
      memberData: { region: "", group: "", country: "" },
      siteUrl: this.props.siteUrl,
      memberrole: "",
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      load: false,
      memberListColumnNames: [],
      configListSettings: [],
      championsList: false,
      isUserAdded: false,
      userStatus: "",
      regionColumnName: "",
      countryColumnName: "",
      groupColumnName: ""
    };

    this.updatePeoplePickerMenuAttributes = this.updatePeoplePickerMenuAttributes.bind(this);
    this.removeButtonEvent = this.removeButtonEvent.bind(this);
    this.onFocusAreaChange = this.onFocusAreaChange.bind(this);
    this.getMemberListColumnNames = this.getMemberListColumnNames.bind(this);
    this.getConfigListSettings = this.getConfigListSettings.bind(this);
    this.populateColumnNames = this.populateColumnNames.bind(this);

    //Create object for CommonServices class
    commonServiceManager = new commonServices(
      this.props.context,
      this.props.siteUrl
    );
  }

  //Get Member list data and config list data into the component
  public async componentDidMount() {
    this.props.context.spHttpClient
      .get(

        "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Region')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((regions: any) => {
          if (!regions.error) {
            this.props.context.spHttpClient
              .get(

                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Country')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line: no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((countries: any) => {
                  if (!countries.error) {
                    this.setState({
                      regions: regions.Choices,
                      countries: countries.Choices,
                    });
                  }
                });
              });
          }
        });
      });

    this.props.context.spHttpClient
      .get(

        "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Group')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((groups: any) => {
          if (!groups.error) {
            this.props.context.spHttpClient
              .get(

                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('FocusArea')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line: no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((focusAreas: any) => {
                  if (!focusAreas.error) {
                    this.setState({
                      groups: groups.Choices,
                      focusAreas: focusAreas.Choices,
                    });
                  }
                });
              });
          }
        });
      });

    //Add Aria required attribute to people picker control
    const inputElement = this.addMemberPeoplePickerParentRef.current.getElementsByTagName("input")[0];
    inputElement.setAttribute('aria-label', 'People picker this is a required field');

    //Update Aria Label attribute to people picker control
    const peoplePickerElement = this.addMemberPeoplePickerParentRef.current.getElementsByClassName('ms-FocusZone');
    const peoplePickerAriaLabel = this.props.isAdmin ? LocaleStrings.AddMemberPageTitle : LocaleStrings.NominateMemberPageTitle;
    peoplePickerElement[0].setAttribute('aria-label', peoplePickerAriaLabel);

    this.updatePeoplePickerMenuAttributes();
    await this.getConfigListSettings();
    await this.getMemberListColumnNames();
  }

  //Get settings from config list
  private async getConfigListSettings() {
    try {
      const configListData: IConfigList[] = await commonServiceManager.getMemberListColumnConfigSettings();
      if (configListData.length === 3) {
        this.setState({ configListSettings: configListData });
      }
      else {
        this.setState({
          errorMessage:
            stringsConstants.CMPErrorMessage +
            ` while loading the page. There could be a problem with the ${stringsConstants.ConfigList} data.`
        });
      }
    }
    catch (error) {
      console.error("CMP_ClbAddMember_getConfigListSettings \n", error);
      this.setState({
        errorMessage:
          stringsConstants.CMPErrorMessage +
          `while retrieving the ${stringsConstants.ConfigList} settings. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //Get memberlist column names from member list
  private async getMemberListColumnNames() {
    try {
      const columnsDisplayNames: any[] = await commonServiceManager.getMemberListColumnDisplayNames();
      if (columnsDisplayNames.length > 0) {
        this.setState({ memberListColumnNames: columnsDisplayNames });
      }
    }
    catch (error) {
      console.error("CMP_AddMember_getMemberListColumnNames \n", error);
      this.setState({
        errorMessage:
          stringsConstants.CMPErrorMessage +
          ` while retrieving the ${stringsConstants.MemberList} column data. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //Update Aria Label attribute to people picker control's suggestions Menu
  private updatePeoplePickerMenuAttributes = () => {
    const inputElement = this.addMemberPeoplePickerParentRef.current.getElementsByTagName('input')[0];

    inputElement.onchange = () => {
      const peopleSuggestions = document.getElementsByClassName('ms-Suggestions-itemButton');

      if (this.addMemberPeoplePickerRef.current.state.mostRecentlyUsedPersons.length > 0 && peopleSuggestions.length > 0) {
        const peoplePickerMenu = document.getElementsByClassName('ms-Suggestions-container')[0];
        peoplePickerMenu.setAttribute("aria-label", this.addMemberPeoplePickerRef.current.props.placeholder);
      }
    };

    const inputEvent = () => {
      setTimeout(() => {
        const peoplePicker = this.addMemberPeoplePickerParentRef.current.getElementsByClassName('ms-FocusZone');
        const peopleSuggestions = document.getElementsByClassName('ms-Suggestions-itemButton');
        if (peoplePicker[0].getAttribute('aria-expanded') === "true" && this.addMemberPeoplePickerRef.current.state.mostRecentlyUsedPersons.length > 0 && peopleSuggestions.length > 0) {
          const peoplePickerMenu = document.getElementsByClassName('ms-Suggestions-container')[0];
          peoplePickerMenu.setAttribute("aria-label", this.addMemberPeoplePickerRef.current.props.placeholder);
        }
      }, 1000);
    };

    inputElement.onclick = inputEvent;
    inputElement.onfocus = inputEvent;

  }

  //onclick event to the people remove button
  private removeButtonEvent = () => {
    if (this.state.UserDetails.length > 0) {
      const removeBtn = this.addMemberPeoplePickerParentRef.current.getElementsByClassName('ms-PickerItem-removeButton')[0];
      if (removeBtn !== undefined) {
        removeBtn.addEventListener('click', () => {
          setTimeout(() => {
            const peoplePicker = this.addMemberPeoplePickerParentRef.current.getElementsByClassName('ms-FocusZone');
            const peopleSuggestions = document.getElementsByClassName('ms-Suggestions-itemButton');
            if (peoplePicker[0].getAttribute('aria-expanded') === "true" && this.addMemberPeoplePickerRef.current.state.mostRecentlyUsedPersons.length > 0 && peopleSuggestions.length > 0) {
              const peoplePickerMenu = document.getElementsByClassName('ms-Suggestions-container')[0];
              peoplePickerMenu.setAttribute("aria-label", this.addMemberPeoplePickerRef.current.props.placeholder);
            }
            this.updatePeoplePickerMenuAttributes();
          }, 1000);
        });
      }
    }
  }

  public componentDidUpdate(prevProps: Readonly<IClbAddMemberProps>, prevState: Readonly<IState>, snapshot?: any): void {
    //Update aria label to suggestions menu when people picker control re-renders on selection
    if (prevState.UserDetails.length !== this.state.UserDetails.length) {
      this.removeButtonEvent();
    }
    if (prevState.multiSelectChoices !== this.state.multiSelectChoices) {
      this.setState({
        selectedFocusAreas: this.state.multiSelectChoices
      });
    }
    //Remove "All" from the array to store it in Members List.
    if (prevState.selectedFocusAreas !== this.state.selectedFocusAreas) {
      let idx = this.state.selectedFocusAreas.indexOf(stringsConstants.AllLabel);
      if (idx != -1)
        this.state.selectedFocusAreas.splice(idx, 1);
    }

    //update dropdown states with member list column display names 
    if (prevState.configListSettings !== this.state.configListSettings ||
      prevState.memberListColumnNames !== this.state.memberListColumnNames) {
      if (this.state.configListSettings.length > 0 && this.state.memberListColumnNames.length > 0)
        this.populateColumnNames();
    }
  }

  //populate member list column display names into the states
  private populateColumnNames() {
    const enabledSettingsArray = this.state.configListSettings.filter((setting) => setting.Value === stringsConstants.EnabledStatus);
    for (let setting of enabledSettingsArray) {
      const columnObject = this.state.memberListColumnNames.find((column) => column.InternalName === setting.Title);
      if (columnObject.InternalName === stringsConstants.RegionColumn) {
        this.setState({ regionColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringsConstants.CountryColumn) {
        this.setState({ countryColumnName: columnObject.Title });
        continue;
      }
      if (columnObject.InternalName === stringsConstants.GroupColumn) {
        this.setState({ groupColumnName: columnObject.Title });
      }
    }
  }

  private _getPeoplePickerItems(items: any[]) {
    let userarr: IUserDetail[] = [];
    items.forEach((user) => {
      userarr.push({ ID: user.id, LoginName: user.loginName });
    });
    this.setState({ UserDetails: userarr });
    if (items.length === 0) this.setState({ updatedMessage: "" });
  }

  private async _getListData(email: any): Promise<any> {
    return this.props.context.spHttpClient
      .get(
        "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/Items?$filter=Title eq '" + email.toLowerCase() + "'",
        SPHttpClient.configurations.v1
      )
      .then(async (response: SPHttpClientResponse) => {
        if (response.status === 200) {
          let flag = 0;
          await response.json().then((responseJSON: any) => {
            let i = 0;
            while (i < responseJSON.value.length) {
              if (
                responseJSON.value[i] &&
                responseJSON.value[i].hasOwnProperty("Title")
              ) {
                if (
                  responseJSON.value[i].Title.toLowerCase() ==
                  email.toLowerCase()
                ) {
                  flag = 1;
                  return flag;
                }
              }
              i++;
            }
            return flag;
          });
          return flag;
        }
      });
  }

  //Add person to the Member List
  public async _createorupdateItem() {
    return this.props.context.spHttpClient
      .get(

        "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          if (!datauser.error) {
            this.props.context.spHttpClient
              .get(

                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/Items",
                SPHttpClient.configurations.v1
              )
              .then((responsen: SPHttpClientResponse) => {
                responsen.json().then((datada: any) => {
                  let memberDataId = datada.value.find(
                    (d: { Title: string }) =>
                      d.Title.toLowerCase() === datauser.Email.toLowerCase()
                  );
                  let memberidData =
                    memberDataId !== undefined
                      ? memberDataId.Role.toLowerCase()
                      : "User";
                  this.setState({ memberrole: memberidData });
                  if (this.state.UserDetails.length > 0) {
                    let email = this.state.UserDetails[0].ID.split("|")[2];
                    // tslint:disable-next-line: no-shadowed-variable
                    this.props.context.spHttpClient
                      .get("/" + this.state.inclusionpath + "/" + this.state.sitename +
                        "/_api/web/siteusers",
                        SPHttpClient.configurations.v1
                      )
                      .then((responseData: SPHttpClientResponse) => {
                        if (responseData.status === 200) {
                          responseData.json().then(async (data: any) => {
                            // tslint:disable-next-line: no-function-expression
                            var member: any = [];
                            data.value.forEach((element: any) => {
                              if (element.Email.toLowerCase() === email.toLowerCase())
                                member.push(element);
                            });
                            const profile = await sp.profiles.getPropertiesFor(this.state.UserDetails[0].LoginName);

                            //get first name and last name from the user profile properties
                            let firstName = "";
                            let lastName = "";
                            for (let i = 0; i < profile.UserProfileProperties.length; i++) {
                              if (firstName === "" || lastName === "") {
                                if (profile.UserProfileProperties[i].Key === "FirstName") {
                                  firstName = profile.UserProfileProperties[i].Value;
                                }
                                if (profile.UserProfileProperties[i].Key === "LastName") {
                                  lastName = profile.UserProfileProperties[i].Value;
                                }
                              }
                              else {
                                break;
                              }
                            }
                            const listDefinition: any = {
                              Title: email,
                              FirstName: firstName,
                              LastName: lastName,
                              Region: this.state.memberData.region,
                              Country: this.state.memberData.country,
                              Role: "Champion",
                              Status:
                                this.state.memberrole === "manager" ||
                                  this.state.memberrole === "Manager" ||
                                  this.state.memberrole === "MANAGER" ||
                                  localStorage["UserRole"] === "Manager"
                                  ? "Approved"
                                  : "Pending",
                              Group: this.state.memberData.group,
                              FocusArea:
                                this.state.selectedFocusAreas.length > 0 ? this.state.selectedFocusAreas : [stringsConstants.TeamWorkLabel],
                            };
                            const spHttpClientOptions: ISPHttpClientOptions = {
                              body: JSON.stringify(listDefinition),
                            };
                            let flag = await this._getListData(email);
                            if (flag == 0) {
                              const url: string =
                                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/items";
                              this.props.context.spHttpClient
                                .post(
                                  url,
                                  SPHttpClient.configurations.v1,
                                  spHttpClientOptions
                                )
                                .then((response: SPHttpClientResponse) => {
                                  if (response.status === 201) {
                                    this.setState({
                                      UserDetails: [],
                                      isAddChampion: false,
                                      load: false
                                    });
                                    this.setState({ isUserAdded: true, userStatus: listDefinition.Status, championsList: true });
                                  } else {
                                    this.setState({
                                      errorMessage: `Response status ${response.status} - ${response.statusText}`,
                                      load: false
                                    });
                                  }
                                });
                            } else {
                              this.setState({
                                updatedMessage: LocaleStrings.UserExistingMessage,
                                load: false
                              });
                            }
                          });
                        } else {
                          this.setState({
                            errorMessage: `Response status ${responseuser.status} - ${responseuser.statusText}`,
                            load: false
                          });
                        }
                      });
                  }
                });
              });
          }
        });
      });
  }

  public filterUsers(type: string, selectedOption: any) {
    if (selectedOption.key !== "All") {
      this.setState({
        memberData: {
          ...this.state.memberData,
          [type]: selectedOption.key,
        },
      });
    }
  }

  //Set state variable whenever the Focus Area dropdown is changed
  public onFocusAreaChange = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): Promise<void> => {
    if (item === undefined) {
      return;
    }
    //Select all the dropdown options when "All" is selected.
    if (item.key === stringsConstants.AllLabel && item.selected) {
      this.setState({
        multiSelectChoices: this.options(this.state.focusAreas).map((option) => option.key as string)
      });
    } //Clear all the dropdown options when "All" is unselected. 
    else if (item.key === stringsConstants.AllLabel) {
      this.setState({ multiSelectChoices: [] });
    } //When an option selected from the dropdown other than "All"
    else if (item.selected) {
      const newKeys = [item.key as string];
      if (this.state.multiSelectChoices.length === this.state.focusAreas.length - 1) {
        newKeys.push(stringsConstants.AllLabel);
      }
      this.setState({ multiSelectChoices: [...this.state.multiSelectChoices, ...newKeys] });
    } //When an option unselected from the dropdown other than "All"
    else {
      this.setState({
        multiSelectChoices: this.state.multiSelectChoices.filter((key: any) => key !== item.key && key !== stringsConstants.AllLabel)
      });
    }
  }

  public options = (optionArray: any) => {
    let myoptions = [];
    if (optionArray !== undefined) {
      myoptions.push({ key: "All", text: "All" });
      optionArray.forEach((element: any) => {
        myoptions.push({ key: element, text: element });
      });
    }
    return myoptions;
  }

  public render() {
    //storing number of dropdowns got enabled
    const enabledDropdownCount = (this.state.countryColumnName !== "" ? 1 : 0) +
      (this.state.regionColumnName !== "" ? 1 : 0) + (this.state.groupColumnName !== "" ? 1 : 0);
    const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
    return (
      <>
        {
          this.state.championsList ? (
            <ClbChampionsList
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              appTitle={this.props.appTitle}
              userAdded={this.state.isUserAdded}
              userStatus={this.state.userStatus}
              onHomeCallBack={this.props.onHomeCallBack}
              configListData={this.state.configListSettings}
              memberListColumnsNames={this.state.memberListColumnNames}
              currentThemeName={this.props.currentThemeName}
            />
          ) :
            <div className='container'>
              <div className={`${styles.addMembersPath}${isDarkOrContrastTheme ? " " + styles.addMembersPathDarkContrast : ""}`}>
                <img src={require("../assets/CMPImages/BackIcon.png")}
                  className={styles.backImg}
                  alt={LocaleStrings.BackButton}
                  aria-hidden="true"
                />
                <span
                  className={styles.backLabel}
                  onClick={() => { this.props.onClickBack(); }}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.props.onClickBack(); }}
                  aria-label={this.props.appTitle}
                >
                  <span title={this.props.appTitle}>
                    {this.props.appTitle}
                  </span>
                </span>
                <span className={styles.border}></span>
                <span className={styles.addMemberLabel}>{this.props.isAdmin ? LocaleStrings.AddMemberPageTitle : LocaleStrings.NominateMemberPageTitle}</span>
              </div>
              {this.state.updatedMessage !== "" ?
                <Label className={`${styles.updatedMessage}${isDarkOrContrastTheme ? " " + styles.updatedMessageDarkContrast : ""}`} aria-live="polite" role="status" id="updateMessagesId">
                  <img src={require('../assets/TOTImages/tickIcon.png')} alt="tickIcon" className={styles.tickImage} />
                  {this.state.updatedMessage}
                </Label> : null}
              {this.state.errorMessage !== "" ?
                <Label className={`${styles.errorMessage}${isDarkOrContrastTheme ? " " + styles.errorMessageDarkContrast : ""}`}
                  aria-live="polite" role="alert">{this.state.errorMessage} </Label> : null}
              <Label className={`${styles.pickerLabel}${isDarkOrContrastTheme ? " " + styles.pickerLabelDarkContrast : ""}`}
                tabIndex={0} aria-label={this.props.isAdmin ? LocaleStrings.AddMemberPageTitle : LocaleStrings.NominateMemberPageTitle}>
                {this.props.isAdmin ? LocaleStrings.AddMemberPageTitle : LocaleStrings.NominateMemberPageTitle} <span className={styles.asterisk}>*</span>
              </Label>
              <div ref={this.addMemberPeoplePickerParentRef}>
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  required={true}
                  onChange={this._getPeoplePickerItems.bind(this)}
                  showHiddenInUI={false}
                  principalTypes={[PrincipalType.User]}
                  defaultSelectedUsers={this.state.selectedusers}
                  resolveDelay={1000}
                  placeholder={LocaleStrings.PeoplePickerPlaceholder}
                  ref={this.addMemberPeoplePickerRef}
                  peoplePickerCntrlclassName={`${styles.addMemberPeoplePickerClass}${isDarkOrContrastTheme ? " " + styles.addMemberPeoplePickerClassDarkContrast : ""}`}
                />
              </div>
              <br></br>
              <Row>
                {this.state.memberListColumnNames.length > 0 ?
                  <>
                    {this.state.configListSettings.length === 3 &&
                      <>
                        {this.state.regionColumnName !== "" &&
                          <Col md={enabledDropdownCount < 3 ? 4 : 3} sm={8}>
                            <span aria-label={LocaleStrings.RegionLabel} className={`${styles.labelContent}${isDarkOrContrastTheme ? " " + styles.labelContentDarkContrast : ""}`}>
                              {LocaleStrings.RegionLabel}</span>
                            <Dropdown
                              onChange={(_event: any, selectedOption: any) => this.filterUsers("region", selectedOption)}
                              options={this.options(this.state.regions)}
                              ariaLabel={`Select ${this.state.regionColumnName}`}
                              className={styles.addMemberDropdown}
                              calloutProps={{ className: "addMemberDropdownCallout" }}
                              onRenderPlaceholder={() => <span title={`Select ${this.state.regionColumnName}`} aria-hidden="true">
                                Select {this.state.regionColumnName}
                              </span>}
                            />
                          </Col>
                        }
                        {this.state.countryColumnName !== "" &&
                          <Col md={enabledDropdownCount < 3 ? 4 : 3} sm={8}>
                            <span aria-label={LocaleStrings.CountryGridHeader} className={`${styles.labelContent}${isDarkOrContrastTheme ? " " + styles.labelContentDarkContrast : ""}`}>
                              {LocaleStrings.CountryGridHeader}</span>
                            <Dropdown
                              onChange={(event: any, selectedOption: any) => this.filterUsers("country", selectedOption)}
                              options={this.options(this.state.countries)}
                              ariaLabel={`Select ${this.state.countryColumnName}`}
                              className={styles.addMemberDropdown}
                              calloutProps={{ className: "addMemberDropdownCallout" }}
                              onRenderPlaceholder={() => <span title={`Select ${this.state.countryColumnName}`} aria-hidden="true">
                                Select {this.state.countryColumnName}
                              </span>}
                            />
                          </Col>
                        }
                        {this.state.groupColumnName !== "" &&
                          <Col md={enabledDropdownCount < 3 ? 4 : 3} sm={8}>
                            <span aria-label={LocaleStrings.GroupGridHeader} className={`${styles.labelContent}${isDarkOrContrastTheme ? " " + styles.labelContentDarkContrast : ""}`}>
                              {LocaleStrings.GroupGridHeader}</span>
                            <Dropdown
                              onChange={(event: any, selectedOption: any) => this.filterUsers("group", selectedOption)}
                              options={this.options(this.state.groups)}
                              ariaLabel={`Select ${this.state.groupColumnName}`}
                              className={styles.addMemberDropdown}
                              calloutProps={{ className: "addMemberDropdownCallout" }}
                              onRenderPlaceholder={() => <span title={`Select ${this.state.groupColumnName}`} aria-hidden="true">
                                Select {this.state.groupColumnName}
                              </span>}
                            />
                          </Col>
                        }
                      </>
                    }
                    <Col md={enabledDropdownCount < 3 ? 4 : 3} sm={8}>
                      <span aria-label={LocaleStrings.FocusAreaLabel} className={`${styles.labelContent}${isDarkOrContrastTheme ? " " + styles.labelContentDarkContrast : ""}`}>
                        {LocaleStrings.FocusAreaLabel}</span>
                      <Dropdown
                        onChange={this.onFocusAreaChange.bind(this)}
                        placeholder={LocaleStrings.FocusAreaPlaceholder}
                        options={this.options(this.state.focusAreas)}
                        ariaLabel={LocaleStrings.FocusAreaPlaceholder}
                        multiSelect
                        selectedKeys={this.state.multiSelectChoices}
                        className={styles.addMemberDropdown}
                        calloutProps={{ className: "addMemberDropdownCallout", doNotLayer: true }}
                        onRenderPlaceholder={() =>
                          <span title={LocaleStrings.FocusAreaPlaceholder} aria-hidden="true">
                            {LocaleStrings.FocusAreaPlaceholder}
                          </span>
                        }
                        onRenderTitle={(options: any) => {
                          const selectedAreas = options.map((option: any) => option.text).join(", ");
                          return (<span aria-hidden="true">{selectedAreas}</span>);
                        }}
                      />
                    </Col>
                  </> : null
                }
              </Row>
              <div className={`${styles.btnArea}${isDarkOrContrastTheme ? " " + styles.btnAreaDarkContrast : ""}`}>
                <button
                  className={`btn ${styles.cancelBtn}`}
                  onClick={() => this.props.onClickBack()}
                  title={LocaleStrings.BackButton}
                >
                  <Icon iconName="NavigateBack" className={`${styles.cancelBtnIcon}`} />
                  <span className={styles.cancelBtnLabel}>{LocaleStrings.BackButton}</span>
                </button>
                <button
                  className={`btn ${styles.saveBtn}`}
                  onClick={() => {
                    this._createorupdateItem();
                    this.state.UserDetails.length > 0 ? this.setState({ load: true }) : this.setState({ load: false });
                  }}
                  title={LocaleStrings.SaveButton}
                  aria-labelledby='updateMessagesId'
                >
                  <Icon iconName="Save" className={`${styles.saveBtnIcon}`} />
                  <span className={styles.saveBtnLabel}>{LocaleStrings.SaveButton}</span>
                </button>
              </div>
              {this.state.load && <div className={styles.load}></div>}
            </div>
        }
      </>
    );
  }
}

export default ClbAddMember;
