import { Checkbox, ResponsiveMode, SearchBox, Spinner, SpinnerSize } from "@fluentui/react";
import { Button } from "@fluentui/react-components";
import {
  Add24Regular,
  ArrowCircleLeft24Regular,
  Save24Regular,
} from "@fluentui/react-icons";
import { Person } from "@microsoft/mgt-react/dist/es6/spfx";
import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import { sp } from "@pnp/sp";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as LocaleStrings from "ClbHomeWebPartStrings";
import { Dropdown, IDropdownOption } from "office-ui-fabric-react/lib/Dropdown";
// import { Dropdown, IDropdownOption } from "@fluentui/react";
import { Label } from "office-ui-fabric-react/lib/Label";
import * as React from "react";
import BootstrapTable from "react-bootstrap-table-next";
import paginationFactory from "react-bootstrap-table2-paginator";
import ToolkitProvider, {
  ToolkitContextType,
} from "react-bootstrap-table2-toolkit";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import commonServices from "../Common/CommonServices";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/CMPAddMember.module.scss";
import ClbChampionsList from "./ClbChampionsList";
import { IConfigList } from "./ManageConfigSettings";

export interface ISPList {
  Title: string;
  FirstName: string;
  LastName: string;
  Country: string;
  Status: string;
  FocusArea: string;
  Group: string;
  Role: string;
  Region: string;
  Points: number;
  ID: number;
}

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
interface IUserDetail {
  ID: any;
  LoginName: string;
}
interface IState {
  errorMessage: string;
  updatedMessage: string;
  UserDetails: Array<IUserDetail>;
  siteUrl: string;
  countries: Array<any>;
  regions: Array<any>;
  groups: Array<any>;
  focusAreas: Array<any>;
  selectedFocusAreas: any;
  multiSelectChoices: any;
  memberData: any;
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
  isChampionAdmin: boolean;
  membersList: any;
  selectedMembers: any;
  showSpinner: boolean;
  isAllSelected: boolean;
  modifiedMembersArray: any;
  isUserUpdated: boolean;
  disableSaveBtn: boolean;
  selectedRegion: any;
  selectedCountry: any;
  selectedGroup: any;
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
      errorMessage: "",
      updatedMessage: "",
      UserDetails: [],
      countries: [],
      regions: [],
      groups: [],
      focusAreas: [],
      selectedFocusAreas: [],
      multiSelectChoices: [],
      memberData: {
        region: "",
        group: "",
        country: "",
        focusArea: "",
        status: "",
        memberName: "",
        isAdmin: false,
      },
      siteUrl: this.props.siteUrl,
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
      groupColumnName: "",
      isChampionAdmin: false,
      membersList: [],
      selectedMembers: [],
      showSpinner: false,
      isAllSelected: false,
      modifiedMembersArray: [],
      isUserUpdated: false,
      disableSaveBtn: false,
      selectedCountry: [],
      selectedGroup: [],
      selectedRegion: [],
    };

    this.updatePeoplePickerMenuAttributes =
      this.updatePeoplePickerMenuAttributes.bind(this);
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
    let regionsOptions = await this.props.context.spHttpClient.get(
      "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('" +
        stringsConstants.MemberList +
        "')/fields/GetByInternalNameOrTitle('Region')",
      SPHttpClient.configurations.v1
    );

    let countriesOptions = await this.props.context.spHttpClient.get(
      "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('" +
        stringsConstants.MemberList +
        "')/fields/GetByInternalNameOrTitle('Country')",
      SPHttpClient.configurations.v1
    );

    let groupsOptions = await this.props.context.spHttpClient.get(
      "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('" +
        stringsConstants.MemberList +
        "')/fields/GetByInternalNameOrTitle('Group')",
      SPHttpClient.configurations.v1
    );

    let focusAreaOptions = await this.props.context.spHttpClient.get(
      "/" +
        this.state.inclusionpath +
        "/" +
        this.state.sitename +
        "/_api/web/lists/GetByTitle('" +
        stringsConstants.MemberList +
        "')/fields/GetByInternalNameOrTitle('FocusArea')",
      SPHttpClient.configurations.v1
    );

    this.setState({
      regions: (await regionsOptions.json()).Choices,
      countries: (await countriesOptions.json()).Choices,
      focusAreas: (await focusAreaOptions.json()).Choices,
      groups: (await groupsOptions.json()).Choices || [],
    });

    //Add Aria required attribute to people picker control
    const inputElement =
      this.addMemberPeoplePickerParentRef.current.getElementsByTagName(
        "input"
      )[0];
    inputElement.setAttribute(
      "aria-label",
      //"People picker this is a required field"
      `Member name ${LocaleStrings.PeoplePickerPlaceholder}`
    );

    // bug 17163
    inputElement.setAttribute("id", "memberName");
   const labelElement = document.getElementsByClassName("ms-BasePicker")[0]?.previousElementSibling;
   labelElement.setAttribute("for","memberName");
   labelElement.setAttribute("aria-label", "Member name");
    // inputElement.setAttribute("aria-labelledby", "addMember");

    //Update Aria Label attribute to people picker control
    const peoplePickerElement =
      this.addMemberPeoplePickerParentRef.current.getElementsByClassName(
        "ms-FocusZone"
      );
    const peoplePickerAriaLabel = this.props.isAdmin
      ? LocaleStrings.AddMemberPageTitle
      : LocaleStrings.NominateMemberPageTitle;
    peoplePickerElement[0].setAttribute("aria-label", peoplePickerAriaLabel);

    this.updatePeoplePickerMenuAttributes();
    await this.getConfigListSettings();
    await this.getMemberListColumnNames();
    await this.getMembersData();
  }

  //Get settings from config list
  private async getConfigListSettings() {
    try {
      const configListData: IConfigList[] =
        await commonServiceManager.getMemberListColumnConfigSettings();
      if (configListData.length === 3) {
        this.setState({ configListSettings: configListData });
      } else {
        this.setState({
          errorMessage:
            stringsConstants.CMPErrorMessage +
            ` while loading the page. There could be a problem with the ${stringsConstants.ConfigList} data.`,
        });
      }
    } catch (error) {
      console.error("CMP_ClbAddMember_getConfigListSettings \n", error);
      this.setState({
        errorMessage:
          stringsConstants.CMPErrorMessage +
          `while retrieving the ${stringsConstants.ConfigList} settings. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //get memebers list data
  private async getMembersData() {
    try {
      //Getting the pending items from Member List
      let filterQuery = "Status eq '" + stringsConstants.approvedStatus + "'";
      const sortColumn = "Role";
      const membersData: any[] = await commonServiceManager.getItemsSortedWithFilter(stringsConstants.MemberList, filterQuery, sortColumn);
      const adminMembers: any[] = membersData.filter((member) => member.Role === stringsConstants.ManagerString || member.Role === stringsConstants.AdminString);

      this.setState({ membersList: membersData, selectedMembers: adminMembers.map((member) => member.ID)});
    } catch (error) {
      console.log(error);
    }
  }

  //Get memberlist column names from member list
  private async getMemberListColumnNames() {
    try {
      const columnsDisplayNames: any[] =
        await commonServiceManager.getMemberListColumnDisplayNames();
      if (columnsDisplayNames.length > 0) {
        this.setState({ memberListColumnNames: columnsDisplayNames });
      }
    } catch (error) {
      console.error("CMP_AddMember_getMemberListColumnNames \n", error);
      this.setState({
        errorMessage:
          stringsConstants.CMPErrorMessage +
          ` while retrieving the ${stringsConstants.MemberList} column data. Below are the details: \n` +
          JSON.stringify(error),
      });
    }
  }

  //Update Aria Label attribute to people picker c  ontrol's suggestions Menu
  private updatePeoplePickerMenuAttributes = () => {
    const inputElement =
      this.addMemberPeoplePickerParentRef.current.getElementsByTagName(
        "input"
      )[0];

    inputElement.onchange = () => {
      const peopleSuggestions = document.getElementsByClassName(
        "ms-Suggestions-itemButton"
      );

      if (
        this.addMemberPeoplePickerRef.current.state.mostRecentlyUsedPersons
          .length > 0 &&
        peopleSuggestions.length > 0
      ) {
        const peoplePickerMenu = document.getElementsByClassName(
          "ms-Suggestions-container"
        )[0];
        peoplePickerMenu.setAttribute(
          "aria-label",
          this.addMemberPeoplePickerRef.current.props.placeholder
        );
      }
    };

    const inputEvent = () => {
      setTimeout(() => {
        const peoplePicker =
          this.addMemberPeoplePickerParentRef.current.getElementsByClassName(
            "ms-FocusZone"
          );
        const peopleSuggestions = document.getElementsByClassName(
          "ms-Suggestions-itemButton"
        );
        if (
          peoplePicker[0].getAttribute("aria-expanded") === "true" &&
          this.addMemberPeoplePickerRef.current.state.mostRecentlyUsedPersons
            .length > 0 &&
          peopleSuggestions.length > 0
        ) {
          const peoplePickerMenu = document.getElementsByClassName(
            "ms-Suggestions-container"
          )[0];
          peoplePickerMenu.setAttribute(
            "aria-label",
            this.addMemberPeoplePickerRef.current.props.placeholder
          );
        }
      }, 1000);
    };

    inputElement.onclick = inputEvent;
    inputElement.onfocus = inputEvent;
  };

  //onclick event to the people remove button
  private removeButtonEvent = () => {
    if (this.state.UserDetails.length > 0) {
      const removeBtn =
        this.addMemberPeoplePickerParentRef.current.getElementsByClassName(
          "ms-PickerItem-removeButton"
        )[0];
      if (removeBtn !== undefined) {
        removeBtn.addEventListener("click", () => {
          setTimeout(() => {
            const peoplePicker =
              this.addMemberPeoplePickerParentRef.current.getElementsByClassName(
                "ms-FocusZone"
              );
            const peopleSuggestions = document.getElementsByClassName(
              "ms-Suggestions-itemButton"
            );
            if (
              peoplePicker[0].getAttribute("aria-expanded") === "true" &&
              this.addMemberPeoplePickerRef.current.state
                .mostRecentlyUsedPersons.length > 0 &&
              peopleSuggestions.length > 0
            ) {
              const peoplePickerMenu = document.getElementsByClassName(
                "ms-Suggestions-container"
              )[0];
              peoplePickerMenu.setAttribute(
                "aria-label",
                this.addMemberPeoplePickerRef.current.props.placeholder
              );
            }
            this.updatePeoplePickerMenuAttributes();
          }, 1000);
        });
      }
    }
  };

  public componentDidUpdate(
    prevProps: Readonly<IClbAddMemberProps>,
    prevState: Readonly<IState>,
    snapshot?: any
  ): void {
    //Update aria label to suggestions menu when people picker control re-renders on selection
    if (prevState.UserDetails.length !== this.state.UserDetails.length) {
      this.removeButtonEvent();
    }
    if (prevState.multiSelectChoices !== this.state.multiSelectChoices) {
      this.setState((prevState) => ({
        selectedFocusAreas: [...prevState.multiSelectChoices],
      }));
    }
    //Remove "All" from the array to store it in Members List.
    if (prevState.selectedFocusAreas !== this.state.selectedFocusAreas) {
      let idx = this.state.selectedFocusAreas.indexOf(
        stringsConstants.AllLabel
      );
      if (idx != -1) this.state.selectedFocusAreas.splice(idx, 1);
    }

    //update dropdown states with member list column display names
    if (
      prevState.configListSettings !== this.state.configListSettings ||
      prevState.memberListColumnNames !== this.state.memberListColumnNames
    ) {
      if (
        this.state.configListSettings.length > 0 &&
        this.state.memberListColumnNames.length > 0
      )
        this.populateColumnNames();
    }
  }

  //populate member list column display names into the states
  private populateColumnNames() {
    const enabledSettingsArray = this.state.configListSettings.filter(
      (setting) => setting.Value === stringsConstants.EnabledStatus
    );
    for (let setting of enabledSettingsArray) {
      const columnObject = this.state.memberListColumnNames.find(
        (column) => column.InternalName === setting.Title
      );
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

  //method to get items from people picker control
  private getPeoplePickerItems(items: any[]) {
    let userarr: IUserDetail[] = [];
    items.forEach((user) => {
      userarr.push({ ID: user.id, LoginName: user.loginName });
    });
    this.setState({
      UserDetails: userarr,
      isUserAdded: false,
      isUserUpdated: false,
    });
    if (items.length === 0) this.setState({ updatedMessage: "" });
  }

  //Check whether User exists in Member List
  private async checkUserInMemberList(email: any): Promise<any> {
    return this.props.context.spHttpClient
      .get(
        "/" +
          this.state.inclusionpath +
          "/" +
          this.state.sitename +
          "/_api/web/lists/GetByTitle('" +
          stringsConstants.MemberList +
          "')/Items?$filter=Title eq '" +
          email.toLowerCase() +
          "'",
        SPHttpClient.configurations.v1
      )
      .then(async (response: SPHttpClientResponse) => {
        if (response.status === 200) {
          let flag = 0;
          await response.json().then((responseJSON: any) => {
            let i = 0;
            while (i < responseJSON.value.length) {
              if (responseJSON?.value[i]?.hasOwnProperty("Title")) {
                if (
                  responseJSON.value[i].Title.toLowerCase() ===
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

  //Add User/Champion to the Member List
  public async createorupdateItem() {
    return (
      this.props.context.spHttpClient
        //Get current logged in user details
        .get(
          "/" +
            this.state.inclusionpath +
            "/" +
            this.state.sitename +
            "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
          SPHttpClient.configurations.v1
        )
        .then((loggedinUserResponse: SPHttpClientResponse) => {
          loggedinUserResponse.json().then((loggedinUserData: any) => {
            if (!loggedinUserData.error) {
              //Update Champion to Member List
              if (this.state.UserDetails.length > 0) {
                let email = this.state.UserDetails[0].ID.split("|")[2];
                this.props.context.spHttpClient
                  .get(
                    "/" +
                      this.state.inclusionpath +
                      "/" +
                      this.state.sitename +
                      "/_api/web/siteusers",
                    SPHttpClient.configurations.v1
                  )
                  .then(async (responseData: SPHttpClientResponse) => {
                    if (responseData.status === 200) {
                      const profile = await sp.profiles.getPropertiesFor(
                        this.state.UserDetails[0].LoginName
                      );

                      //get first name and last name from the user profile properties
                      let firstName = "";
                      let lastName = "";
                      for (let property of profile.UserProfileProperties) {
                        if (firstName === "" || lastName === "") {
                          if (property.Key === "FirstName") {
                            firstName = property.Value;
                          }
                          if (property.Key === "LastName") {
                            lastName = property.Value;
                          }
                        } else {
                          break;
                        }
                      }

                      //Set Member object to add to Member List
                      const listDefinition: any = {
                        Title: email,
                        FirstName: firstName,
                        LastName: lastName,
                        Region: this.state.memberData.region,
                        Country: this.state.memberData.country,
                        Role:
                          this.props.isAdmin && this.state.isChampionAdmin
                            ? stringsConstants.AdminString
                            : stringsConstants.ChampionString,
                        Status: this.props.isAdmin
                          ? stringsConstants.approvedStatus
                          : stringsConstants.pendingStatus,
                        Group: this.state.memberData.group,
                        FocusArea:
                          this.state.selectedFocusAreas.length > 0
                            ? this.state.selectedFocusAreas
                            : [stringsConstants.TeamWorkLabel],
                      };

                      //Set HTTP Post request options
                      const spHttpClientOptions: ISPHttpClientOptions = {
                        body: JSON.stringify(listDefinition),
                      };
                      //check whether the user already exists in the member list
                      let flag = await this.checkUserInMemberList(email);
                      //Add User to Member List
                      if (flag == 0) {
                        const url: string =
                          "/" +
                          this.state.inclusionpath +
                          "/" +
                          this.state.sitename +
                          "/_api/web/lists/GetByTitle('" +
                          stringsConstants.MemberList +
                          "')/items";
                        this.props.context.spHttpClient
                          .post(
                            url,
                            SPHttpClient.configurations.v1,
                            spHttpClientOptions
                          )
                          .then((response: SPHttpClientResponse) => {
                            if (response.status === 201) {
                              this.addMemberPeoplePickerRef.current.setState({
                                selectedPersons: [],
                              });
                              this.addMemberPeoplePickerParentRef.current
                                .getElementsByTagName("input")[0]
                                .blur();
                              if (this.props.isAdmin) {
                                this.getMembersData();
                              }
                              this.setState({
                                UserDetails: [],
                                load: false,
                                isUserAdded: true,
                                userStatus: listDefinition.Status,
                                championsList:
                                  listDefinition.Status ===
                                  stringsConstants.pendingStatus,
                                isChampionAdmin: false,
                                multiSelectChoices: [],
                                selectedCountry: [{ key: "All", text: "All" }],
                                selectedGroup: [{ key: "All", text: "All" }],
                                selectedRegion: [{ key: "All", text: "All" }],
                              });
                            } else {
                              this.setState({
                                errorMessage: `Response status ${response.status} - ${response.statusText}`,
                                load: false,
                              });
                            }
                          });
                      } else {
                        this.setState({
                          updatedMessage: LocaleStrings.UserExistingMessage,
                          load: false,
                        });
                      }
                    } else {
                      this.setState({
                        errorMessage: `Response status ${loggedinUserResponse.status} - ${loggedinUserResponse.statusText}`,
                        load: false,
                      });
                    }
                  });
              }
            }
          });
        })
    );
  }

  //Update Modified Members to Member List
  private async updateMembersItem() {
    try {
      this.setState({ load: true });
      let membersItemArray: any[] = [];
      this.state.memberData.map((member: any) => {
        if (
          this.state.selectedMembers.filter(
            (memberId: number) => memberId !== member.ID
          )
        ) {
          const listDefinition: any = {
            id: member.ID,
            value: {
              Role: member.Role,
            },
          };
          membersItemArray.push(listDefinition);
        }
      });
      await commonServiceManager.updateMultipleItemsWithDifferentValues(stringsConstants.MemberList, membersItemArray).then((data) => {
        this.setState({ load: false, isUserUpdated: true });
        this.getMembersData();
      }).catch((error) => { console.log(error) });
    } catch (error) {
      console.log(error);
    }
  }

  //Filter members based on the dropdown selection
  public filterUsers(type: string, selectedOption: any) {
    //set default values to the dropdowns
    type === "region"
      ? this.setState({ selectedRegion: selectedOption.key })
      : type === "country"
      ? this.setState({ selectedCountry: selectedOption.key })
      : this.setState({ selectedGroup: selectedOption.key });

    if (selectedOption.key !== "All") {
      this.setState((prevState) => ({
        memberData: {
          ...prevState.memberData,
          [type]: selectedOption.key,
        },
      }));
    }
  }

  //Set state variable whenever the Focus Area dropdown is changed
  public onFocusAreaChange = async (
    event: React.FormEvent<HTMLDivElement>,
    item: IDropdownOption
  ): Promise<void> => {
    if (item === undefined) {
      return;
    }
    //Select all the dropdown options when "All" is selected.
    if (item.key === stringsConstants.AllLabel && item.selected) {
      let selectedChoices = this.options(this.state.focusAreas).map(
        (option) => option.key as string
      );
      this.setState({
        multiSelectChoices: selectedChoices,
      });
    } //Clear all the dropdown options when "All" is unselected.
    else if (item.key === stringsConstants.AllLabel) {
      this.setState({ multiSelectChoices: [] });
    } //When an option selected from the dropdown other than "All"
    else if (item.selected) {
      const newKeys = [item.key as string];
      if (
        this.state.multiSelectChoices.length ===
        this.state.focusAreas.length - 1
      ) {
        newKeys.push(stringsConstants.AllLabel);
      }
      this.setState((prevState) => ({
        multiSelectChoices: [...prevState.multiSelectChoices, ...newKeys],
      }));
    } //When an option unselected from the dropdown other than "All"
    else {
      this.setState((prevState) => ({
        multiSelectChoices: prevState.multiSelectChoices.filter(
          (key: any) => key !== item.key && key !== stringsConstants.AllLabel
        ),
      }));
    }
  };

  public options = (optionArray: any) => {
    let myoptions = [];
    if (optionArray !== undefined) {
      myoptions.push({ key: "All", text: "All" });
      optionArray.forEach((element: any) => {
        myoptions.push({ key: element, text: element });
      });
    }
    return myoptions;
  };

  // format the cell for Champion Name
  public championFormatter = (
    cell: any,
    gridRow: any,
    rowIndex: any,
    formatExtraData: any
  ) => {
    return (
      <Person
        personQuery={gridRow.Title}
        view={3}
        personCardInteraction={1}
        className="champion-person-card"
      />
    );
  };

  //render the sort caret on the header column for accessbility
  public customSortCaret = (order: any, column: any) => {
    if (!order) {
      return (
        <span className="sort-order">
          <span className="dropdown-caret"></span>
          <span className="dropup-caret"></span>
        </span>
      );
    } else if (order === "asc") {
      return (
        <span className="sort-order">
          <span className="dropup-caret"></span>
        </span>
      );
    } else if (order === "desc") {
      return (
        <span className="sort-order">
          <span className="dropdown-caret"></span>
        </span>
      );
    }
    return null;
  };

  //Set pagination properties
  private pagination = paginationFactory({
    page: 1,
    sizePerPage: 10,
    lastPageText: ">>",
    firstPageText: "<<",
    nextPageText: ">",
    prePageText: "<",
    showTotal: true,
    alwaysShowAllBtns: false,
    //customized the render options for pagesize button in the pagination for accessbility
    sizePerPageRenderer: ({
      options,
      currSizePerPage,
      onSizePerPageChange,
    }) => (
      <div className="btn-group" role="group">
        {options.map((option) => {
          const isSelect = currSizePerPage === `${option.page}`;
          return (
            <button
              key={option.text}
              type="button"
              onClick={() => onSizePerPageChange(option.page)}
              className={`btn${
                isSelect ? " sizeperpage-selected" : " sizeperpage"
              }${
                this.props.currentThemeName ===
                stringsConstants.themeDefaultMode
                  ? ""
                  : " selected-darkcontrast"
              }`}
              aria-label={
                isSelect
                  ? stringsConstants.sizePerPageLabel +
                    option.text +
                    stringsConstants.selectedAriaLabel
                  : stringsConstants.sizePerPageLabel + option.text
              }
            >
              {option.text}
            </button>
          );
        })}
      </div>
    ),
    //customized the render options for page list in the pagination for accessbility
    pageButtonRenderer: (options: any) => {
      const handleClick = (e: any) => {
        e.preventDefault();
        if (options.disabled) return;
        options.onPageChange(options.page);
      };
      const className = `${options.active ? "active " : ""}${
        options.disabled ? "disabled " : ""
      }`;
      let ariaLabel = "";
      let pageText = "";
      switch (options.title) {
        case "first page":
          ariaLabel = `Go to ${options.title}`;
          pageText = "<<";
          break;
        case "previous page":
          ariaLabel = `Go to ${options.title}`;
          pageText = "<";
          break;
        case "next page":
          ariaLabel = `Go to ${options.title}`;
          pageText = ">";
          break;
        case "last page":
          ariaLabel = `Go to ${options.title}`;
          pageText = ">>";
          break;
        default:
          ariaLabel = `Go to page ${options.title}`;
          pageText = options.title;
          break;
      }
      return (
        <li
          key={options.title}
          className={`${className}page-item${
            this.props.currentThemeName === stringsConstants.themeDefaultMode
              ? ""
              : " selected-darkcontrast"
          }`}
          role="presentation"
          title={ariaLabel}
        >
          <a
            className="page-link"
            href="#"
            onClick={handleClick}
            role="button"
            aria-label={
              options.active
                ? ariaLabel + stringsConstants.selectedAriaLabel
                : ariaLabel
            }
          >
            <span aria-hidden="true">{pageText}</span>
          </a>
        </li>
      );
    },
    paginationTotalRenderer: (from: any, to: any, size: any) => {
      const resultsFound =
        size !== 0 ? `Showing ${from} to ${to} of ${size} Results` : "";
      return (
        <span
          className="react-bootstrap-table-pagination-total"
          aria-live="polite"
          role="status"
        >
          &nbsp;{resultsFound}
        </span>
      );
    },
  });

  //Get Table Header Class
  private getTableHeaderClass(enabledColumnCount: number) {
    switch (enabledColumnCount) {
      case 3: {
        return styles.adminsApprovalTableHeaderWithAllCols;
      }
      case 2: {
        return styles.adminsApprovalTableHeaderWithSixCols;
      }
      case 1: {
        return styles.adminsApprovalTableHeaderWithFiveCols;
      }
      case 0: {
        return styles.adminsApprovalTableHeaderWithFourCols;
      }
    }
  }

  //Get Table Body Class
  private getTableBodyClass(enabledColumnCount: number) {
    switch (enabledColumnCount) {
      case 3: {
        return styles.adminsApprovalTableBodyWithAllCols;
      }
      case 2: {
        return styles.adminsApprovalTableBodyWithSixCols;
      }
      case 1: {
        return styles.adminsApprovalTableBodyWithFiveCols;
      }
      case 0: {
        return styles.adminsApprovalTableBodyWithFourCols;
      }
    }
  }

  //Update all the selected members to Member List
  public selectMembers(isChecked: boolean, key: number) {
    const updateMemberData = this.state.memberData.map((member: any) => {
      if (member.ID === key) {
        member.Role = isChecked
          ? stringsConstants.AdminString
          : stringsConstants.ChampionString;
        member.Status = isChecked
          ? stringsConstants.approvedStatus
          : stringsConstants.pendingStatus;
      }
      return member;
    });
    let disableSaveBtn = !this.state.membersList.some((member: any) => (member.Role === stringsConstants.ManagerString || member.Role === stringsConstants.AdminString));
    if (isChecked) {
      this.setState((prevState) => ({
        selectedMembers: [...prevState.selectedMembers, key],
        memberData: updateMemberData,
        disableSaveBtn: false,
      }));
    } else {
      this.setState((prevState) => ({
        selectedMembers: prevState.selectedMembers.filter(
          (member: number) => member !== key
        ),
        memberData: updateMemberData,
        disableSaveBtn: disableSaveBtn,
      }));
    }
  }

  public render() {
    //check whether the current theme is dark or contrast
    const isDarkOrContrastTheme =
      this.props.currentThemeName === stringsConstants.themeDarkMode ||
      this.props.currentThemeName === stringsConstants.themeContrastMode;

    //storing number of dropdowns got enabled
    const enabledColumnCount =
      (this.state.countryColumnName !== "" ? 1 : 0) +
      (this.state.regionColumnName !== "" ? 1 : 0) +
      (this.state.groupColumnName !== "" ? 1 : 0);

    const adminsTableHeader = [
      {
        dataField: "FirstName",
        text: LocaleStrings.PeopleNameGridHeader,
        headerTitle: true,
        formatter: this.championFormatter,
        searchable: true,
        sort: true,
        sortCaret: this.customSortCaret,
      },
      {
        dataField: "Region",
        text: this.state.regionColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.regionColumnName === "",
      },
      {
        dataField: "Country",
        text: this.state.countryColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.countryColumnName === "",
      },
      {
        dataField: "Group",
        text: this.state.groupColumnName,
        headerTitle: true,
        title: true,
        searchable: false,
        hidden: this.state.groupColumnName === "",
      },
      {
        dataField: "FocusArea",
        text: LocaleStrings.FocusAreaGridHeader,
        headerTitle: true,
        title: true,
        searchable: false,
      },
      {
        dataField: "ID",
        headerTitle: true,
        text: LocaleStrings.AdminLabel,
        title: true,
        attrs: (_cell: any, row: any) => ({ key: row.ID }),
        formatter: (_: any, gridRow: any) => {
          return (
            <Checkbox
              onChange={(_eve: any, isChecked: boolean) => {
                this.selectMembers(isChecked, gridRow.ID);
              }}
              className={styles.selectItemCheckbox}
              checked={this.state.selectedMembers.includes(gridRow.ID)}
              ariaLabel={LocaleStrings.AdminLabel}
              disabled={this.state.showSpinner}
            />
          );
        },
        searchable: false,
      },
    ];

    return (
      <>
        {this.state.championsList ? (
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
        ) : (
          <div className="container">
            <div
              className={`${styles.addMembersPath}${
                isDarkOrContrastTheme
                  ? " " + styles.addMembersPathDarkContrast
                  : ""
              }`}
            >
              <img
                src={require("../assets/CMPImages/BackIcon.png")}
                className={styles.backImg}
                alt={LocaleStrings.BackButton}
                aria-hidden="true"
              />
              <span
                className={styles.backLabel}
                onClick={() => {
                  this.props.onClickBack();
                }}
                role="button"
                tabIndex={0}
                onKeyDown={(evt: any) => {
                  if (evt.key === stringsConstants.stringEnter)
                    this.props.onClickBack();
                }}
                aria-label={this.props.appTitle}
              >
                <span title={this.props.appTitle}>{this.props.appTitle}</span>
              </span>
              <span className={styles.border}></span>
              <span className={styles.addMemberLabel}>
                {this.props.isAdmin
                  ? LocaleStrings.AddMemberPageTitle
                  : LocaleStrings.NominateMemberPageTitle}
              </span>
            </div>

            {this.state.isUserAdded && this.props.isAdmin ? (
              <Label
                className={`${styles.successMessage}${
                  isDarkOrContrastTheme
                    ? " " + styles.successMessageDarkContrast
                    : ""
                }`}
                aria-live="polite"
                role="alert"
                id="successMessagesId"
              >
                <img
                  src={require("../assets/TOTImages/tickIcon.png")}
                  alt="tickIcon"
                  className={styles.tickImage}
                />
                {LocaleStrings.UserAddedMessage}
              </Label>
            ) : null}

            {this.state.errorMessage !== "" ? (
              <Label
                className={`${styles.errorMessage}${
                  isDarkOrContrastTheme
                    ? " " + styles.errorMessageDarkContrast
                    : ""
                }`}
                aria-live="polite"
                role="alert"
              >
                {this.state.errorMessage}{" "}
              </Label>
            ) : null}

            <Label
              className={`${styles.pickerLabel}${
                isDarkOrContrastTheme
                  ? " " + styles.pickerLabelDarkContrast
                  : ""
              }`}
              aria-label={
                this.props.isAdmin
                  ? LocaleStrings.AddMemberPageTitle
                  : LocaleStrings.NominateMemberPageTitle
              }
            >
              {this.props.isAdmin
                ? LocaleStrings.AddMemberPageTitle
                : LocaleStrings.NominateMemberPageTitle}
            </Label>
            <Row xl={2} lg={2} md={2} sm={1} xs={1}>
              <Col
                xl={6}
                lg={6}
                md={8}
                sm={12}
                xs={12}
                ref={this.addMemberPeoplePickerParentRef}
              >
                <PeoplePicker
                  context={this.props.context}
                  personSelectionLimit={1}
                  required={true}
                  onChange={this.getPeoplePickerItems.bind(this)}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                  placeholder={LocaleStrings.PeoplePickerPlaceholder}
                  ref={this.addMemberPeoplePickerRef}
                  titleText="Member Name"
                  peoplePickerWPclassName={`${
                    styles.addMemberPeoplePickerWPClass
                  }${
                    isDarkOrContrastTheme
                      ? " " + styles.addMemberPeoplePickerWPClassDarkContrast
                      : ""
                  }`}
                  peoplePickerCntrlclassName={styles.addMemberPeoplePickerClass}
                />
              </Col>
              {this.props.isAdmin && (
                <Col xl={6} lg={6} md={4} sm={12} xs={12}>
                  <Checkbox
                    label={LocaleStrings.AddMemberAsAdminLabel}
                    className={styles.addAdminCheckbox}
                    onChange={(_, data: any) => {
                      this.setState({ isChampionAdmin: data });
                    }}
                    checked={this.state.isChampionAdmin}
                  />
                </Col>
              )}
            </Row>
            <br></br>
            <Row className="dropDownsWrapper">
              {this.state.memberListColumnNames.length > 0 ? (
                <>
                  {this.state.configListSettings.length === 3 && (
                    <>
                      {this.state.regionColumnName !== "" && (
                        <Col xl={4} lg={4} md={6} sm={12} xs={12}>
                          <div className={styles.addMemberDropdownWP}>
                            <span
                              aria-label={LocaleStrings.RegionLabel}
                              className={`${styles.labelContent}${
                                isDarkOrContrastTheme
                                  ? " " + styles.labelContentDarkContrast
                                  : ""
                              }`}
                            >
                              {LocaleStrings.RegionLabel}
                            </span>
                            <Dropdown
                              onChange={(_event: any, selectedOption: any) =>
                                this.filterUsers("region", selectedOption)
                              }
                              options={this.options(this.state.regions)}
                              selectedKey={this.state.selectedRegion}
                              ariaLabel={`Select ${this.state.regionColumnName}`}
                              // className={styles.addMemberDropdown}
                              className={`${styles.addMemberDropdown}${
                                isDarkOrContrastTheme
                                  ? " " + styles.addMemberDropdownDarkorContrast
                                  : ""
                              }`}
                              calloutProps={{
                                className: "addMemberDropdownCallout",
                              }}
                              onRenderPlaceholder={() => (
                                <span
                                  title={`Select ${this.state.regionColumnName}`}
                                  aria-hidden="true"
                                  className={`${styles.addMemberDropdown}${
                                    isDarkOrContrastTheme
                                      ? " " +
                                        styles.addMemberDropdownDarkorContrast
                                      : ""
                                  }`}
                                >
                                  Select {this.state.regionColumnName}
                                </span>
                              )}
                              responsiveMode={ResponsiveMode.unknown}
                            />
                          </div>
                        </Col>
                      )}
                      {this.state.countryColumnName !== "" && (
                        <Col xl={4} lg={4} md={6} sm={12} xs={12}>
                          <div className={styles.addMemberDropdownWP}>
                            <span
                              aria-label={LocaleStrings.CountryGridHeader}
                              className={`${styles.labelContent}${
                                isDarkOrContrastTheme
                                  ? " " + styles.labelContentDarkContrast
                                  : ""
                              }`}
                            >
                              {LocaleStrings.CountryGridHeader}
                            </span>
                            <Dropdown
                              onChange={(event: any, selectedOption: any) =>
                                this.filterUsers("country", selectedOption)
                              }
                              options={this.options(this.state.countries)}
                              selectedKey={this.state.selectedCountry}
                              ariaLabel={`Select ${this.state.countryColumnName}`}
                              //className={styles.addMemberDropdown}
                              className={`${styles.addMemberDropdown}${
                                isDarkOrContrastTheme
                                  ? " " + styles.addMemberDropdownDarkorContrast
                                  : ""
                              }`}
                              calloutProps={{
                                className: "addMemberDropdownCallout",
                              }}
                              onRenderPlaceholder={() => (
                                <span
                                  title={`Select ${this.state.countryColumnName}`}
                                  aria-hidden="true"
                                  className={`${styles.addMemberDropdown}${
                                    isDarkOrContrastTheme
                                      ? " " +
                                        styles.addMemberDropdownDarkorContrast
                                      : ""
                                  }`}
                                >
                                  Select {this.state.countryColumnName}
                                </span>
                              )}
                              responsiveMode={ResponsiveMode.unknown}
                            />
                          </div>
                        </Col>
                      )}
                      {this.state.groupColumnName !== "" && (
                        <Col xl={4} lg={4} md={6} sm={12} xs={12}>
                          <div className={styles.addMemberDropdownWP}>
                            <span
                              aria-label={LocaleStrings.GroupGridHeader}
                              className={`${styles.labelContent}${
                                isDarkOrContrastTheme
                                  ? " " + styles.labelContentDarkContrast
                                  : ""
                              }`}
                            >
                              {LocaleStrings.GroupGridHeader}
                            </span>
                            <Dropdown
                              onChange={(event: any, selectedOption: any) =>
                                this.filterUsers("group", selectedOption)
                              }
                              options={this.options(this.state.groups)}
                              selectedKey={this.state.selectedGroup}
                              ariaLabel={`Select ${this.state.groupColumnName}`}
                              //className={styles.addMemberDropdown}
                              className={`${styles.addMemberDropdown}${
                                isDarkOrContrastTheme
                                  ? " " + styles.addMemberDropdownDarkorContrast
                                  : ""
                              }`}
                              calloutProps={{
                                className: "addMemberDropdownCallout",
                              }}
                              onRenderPlaceholder={() => (
                                <span
                                  title={`Select ${this.state.groupColumnName}`}
                                  aria-hidden="true"
                                  className={`${styles.addMemberDropdown}${
                                    isDarkOrContrastTheme
                                      ? " " +
                                        styles.addMemberDropdownDarkorContrast
                                      : ""
                                  }`}
                                >
                                  Select {this.state.groupColumnName}
                                </span>
                              )}
                              responsiveMode={ResponsiveMode.unknown}
                            />
                          </div>
                        </Col>
                      )}
                    </>
                  )}
                  <Col xl={4} lg={4} md={6} sm={12} xs={12}>
                    <div className={styles.addMemberDropdownWP}>
                      <span
                        aria-label={LocaleStrings.FocusAreaLabel}
                        className={`${styles.labelContent}${
                          isDarkOrContrastTheme
                            ? " " + styles.labelContentDarkContrast
                            : ""
                        }`}
                      >
                        {LocaleStrings.FocusAreaLabel}
                      </span>
                      <Dropdown
                        onChange={this.onFocusAreaChange.bind(this)}
                        //placeholder={LocaleStrings.FocusAreaPlaceholder}
                        options={this.options(this.state.focusAreas)}
                        ariaLabel={LocaleStrings.FocusAreaPlaceholder}
                        multiSelect
                        selectedKeys={this.state.multiSelectChoices}
                        //className={styles.addMemberDropdown}
                        className={`${styles.addMemberDropdown}${
                          isDarkOrContrastTheme
                            ? " " + styles.addMemberDropdownDarkorContrast
                            : ""
                        }`}
                        calloutProps={{
                          className: "addMemberDropdownCallout",
                          doNotLayer: true
                        }}
                        onRenderPlaceholder={() => (
                          <span
                            title={LocaleStrings.FocusAreaPlaceholder}
                            aria-hidden="true"
                          >
                            {LocaleStrings.FocusAreaPlaceholder}
                          </span>
                        )}
                        onRenderTitle={(options: any) => {
                          const selectedAreas = options
                            .map((option: any) => option.text)
                            .join(", ");
                          return (
                            <span aria-hidden="true">{selectedAreas}</span>
                          );
                        }}
                        responsiveMode={ResponsiveMode.unknown}
                      />
                    </div>
                  </Col>
                </>
              ) : null}
            </Row>
            <div
              className={`${styles.btnArea}${
                isDarkOrContrastTheme ? " " + styles.btnAreaDarkContrast : ""
              }`}
            >
              {!this.props.isAdmin && (
                <Button
                  icon={<ArrowCircleLeft24Regular />}
                  onClick={() => this.props.onClickBack()}
                  onKeyDown={(evt: any) => {
                    if (evt.key === stringsConstants.stringEnter)
                      this.props.onClickBack();
                  }}
                  title={LocaleStrings.BackButton}
                  className={styles.cancelBtn}
                >
                  {LocaleStrings.BackButton}
                </Button>
              )}
              <Button
                icon={<Add24Regular />}
                onClick={() => {
                  this.createorupdateItem();
                  this.state.UserDetails.length > 0
                    ? this.setState({ load: true })
                    : this.setState({ load: false });
                }}
                onKeyDown={(evt: any) => {
                  if (evt.key === stringsConstants.stringEnter) {
                    this.createorupdateItem();
                    this.state.UserDetails.length > 0
                      ? this.setState({ load: true })
                      : this.setState({ load: false });
                  }
                }}
                title={LocaleStrings.SaveButton}
                aria-labelledby="updateMessagesId successMessagesId"
                className={styles.addMemberBtn}
              >
                {this.props.isAdmin
                  ? LocaleStrings.AddMemberPageTitle
                  : LocaleStrings.NominateMemberPageTitle}
              </Button>
            </div>
            {this.state.load && <div className={styles.load}></div>}
            {this.state.isUserUpdated ? (
              <Label
                className={`${styles.updatedMessage}${
                  isDarkOrContrastTheme
                    ? " " + styles.updatedMessageDarkContrast
                    : ""
                }`}
                aria-live="polite"
                role="status"
                id="updateMessagesId"
              >
                <img
                  src={require("../assets/TOTImages/tickIcon.png")}
                  alt="tickIcon"
                  className={styles.tickImage}
                />
                {LocaleStrings.MembersUpdatedMessage}
              </Label>
            ) : null}
            {this.props.isAdmin && (
              <div className="championsListWrapper">
                <ToolkitProvider
                  bootstrap4
                  keyField="ID"
                  data={this.state.membersList}
                  columns={adminsTableHeader}
                  search={{
                    afterSearch: (newResult: ISPList[]) => {
                      this.setState({
                        memberData: newResult,
                        isAllSelected:
                          newResult.length === this.state.selectedMembers.length
                            ? true
                            : false,
                      });
                    },
                  }}
                >
                  {(props: ToolkitContextType) => (
                    <div>
                      <Row xl={2} lg={2} md={2} sm={1} xs={1}>
                        <Col xl={6} lg={6} md={8} sm={12} xs={12}>
                          <label className={styles.tableLabel}>
                            Champions List
                          </label>
                        </Col>
                        <Col xl={6} lg={6} md={4} sm={12} xs={12}>
                          <SearchBox
                            placeholder={
                              LocaleStrings.AddAdminSearchboxPlaceholder
                            }
                            onChange={(_, searchedText) =>
                              props.searchProps.onSearch(searchedText)
                            }
                            className={styles.approvalsSearchbox}
                          />
                        </Col>
                      </Row>
                      <div
                        className={`${styles.approvalsTableContainer}${
                          isDarkOrContrastTheme
                            ? " " + styles.approvalsTableContainerDarkContrast
                            : ""
                        }`}
                      >
                        <BootstrapTable
                          striped
                          {...props.baseProps}
                          table-responsive={true}
                          pagination={this.pagination}
                          wrapperClasses={styles.approvalsTableWrapper}
                          headerClasses={this.getTableHeaderClass(
                            enabledColumnCount
                          )}
                          bodyClasses={this.getTableBodyClass(
                            enabledColumnCount
                          )}
                          noDataIndication={() => (
                            <div className={styles.noRecordsArea}>
                              {this.state.showSpinner ? (
                                <Spinner
                                  label={LocaleStrings.ProcessingSpinnerLabel}
                                  size={SpinnerSize.large}
                                />
                              ) : (
                                <>
                                  <img
                                    src={require("../assets/CMPImages/Norecordsicon.svg")}
                                    alt={LocaleStrings.NoRecordsIcon}
                                    className={styles.noRecordsImg}
                                    aria-hidden={true}
                                  />
                                  <span
                                    className={styles.noRecordsLabels}
                                    aria-live="polite"
                                    role="alert"
                                    tabIndex={0}
                                  >
                                    {this.state.memberData.length === 0
                                      ? LocaleStrings.NoChampionsMessage
                                      : LocaleStrings.NoSearchResults}
                                  </span>
                                </>
                              )}
                            </div>
                          )}
                        />
                      </div>
                    </div>
                  )}
                </ToolkitProvider>
                <div
                  className={`${styles.btnArea}${
                    isDarkOrContrastTheme
                      ? " " + styles.btnAreaDarkContrast
                      : ""
                  }`}
                >
                  <Button
                    icon={<ArrowCircleLeft24Regular />}
                    onClick={() => this.props.onClickBack()}
                    onKeyDown={(evt: any) => {
                      if (evt.key === stringsConstants.stringEnter)
                        this.props.onClickBack();
                    }}
                    title={LocaleStrings.BackButton}
                    className={styles.cancelBtn}
                    tabIndex={0}
                  >
                    {LocaleStrings.BackButton}
                  </Button>
                  <Button
                    icon={<Save24Regular />}
                    className={styles.saveBtn}
                    onClick={() => {
                      this.updateMembersItem();
                      this.state.selectedMembers.length > 0
                        ? this.setState({ load: true })
                        : this.setState({ load: false });
                    }}
                    title={
                      this.state.disableSaveBtn
                        ? LocaleStrings.NoAdminSelectedMessage
                        : LocaleStrings.SaveChangesButton
                    }
                    disabled={this.state.disableSaveBtn}
                    appearance="primary"
                    tabIndex={0}
                  >
                    {LocaleStrings.SaveChangesButton}
                  </Button>
                </div>
              </div>
            )}
          </div>
        )}
      </>
    );
  }
}

export default ClbAddMember;
