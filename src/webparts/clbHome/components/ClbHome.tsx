import {
  ISPHttpClientOptions, SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import "@pnp/sp/lists";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import { Web } from "@pnp/sp/webs";
import { initializeIcons } from "@uifabric/icons";
import {
  teamsLightTheme, teamsDarkTheme, teamsHighContrastTheme,
  FluentProvider, Theme
} from '@fluentui/react-components';
import { app } from "@microsoft/teams-js-v2";
import "bootstrap/dist/css/bootstrap.min.css";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as React from "react";
import Col from "react-bootstrap/Col";
import Media from "react-bootstrap/Media";
import Row from "react-bootstrap/Row";
import siteconfig from "../config/siteconfig.json";
import * as stringsConstants from "../constants/strings";
import styles from "../scss/CMPHome.module.scss";
import ChampionLeaderBoard from "./ChampionLeaderBoard";
import ClbAddMember from "./ClbAddMember";
import DigitalBadge from "./DigitalBadge";
import EmployeeView from "./EmployeeView";
import Header from "./Header";
import { IClbHomeProps } from "./IClbHomeProps";
import ManageApprovals from "./ManageApprovals";
import TOTLandingPage from "./TOTLandingPage";
import ChampionReport from "./ChampionReport";
import CommonServices from "../Common/CommonServices";

initializeIcons();

export interface IClbHomeState {
  cB: boolean;
  clB: boolean;
  addMember: boolean;
  manageApprovals: boolean;
  ChampionsList: boolean;
  cV: boolean;
  siteUrl: string;
  eV: boolean;
  dB: boolean;
  sitename: string;
  inclusionpath: string;
  siteId: any;
  isShow: boolean;
  loggedinUserName: string;
  loggedinUserEmail: string;
  enableTOT: boolean;
  isTOTEnabled: boolean;
  firstName: string;
  appLogoURL: string;
  isChampionApprovalPending: boolean;
  isEventApprovalPending: boolean;
  championReport: boolean;
  appTitle: string;
  currentTeamsTheme: Theme;
  currentThemeName: string;
}

//Global Variables
let flagCheckUserRole: boolean = true;
let rootSiteURL: string;
let spweb: any;
let commonService: CommonServices;
export default class ClbHome extends React.Component<IClbHomeProps, IClbHomeState> {

  constructor(_props: any) {
    super(_props);
    this.state = {
      siteUrl: this.props.siteUrl,
      cB: false,
      addMember: false,
      manageApprovals: false,
      ChampionsList: false,
      clB: false,
      cV: false,
      eV: false,
      dB: false,
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      siteId: "",
      isShow: false,
      loggedinUserName: "",
      loggedinUserEmail: "",
      enableTOT: false,
      isTOTEnabled: false,
      firstName: "",
      appLogoURL: "",
      isChampionApprovalPending: false,
      isEventApprovalPending: false,
      championReport: false,
      appTitle: "",
      currentTeamsTheme: teamsLightTheme,
      currentThemeName: stringsConstants.themeDefaultMode
    };
    this.checkUserRole = this.checkUserRole.bind(this);
    this.callBackFunction = this.callBackFunction.bind(this);
    //set context
    if (this.props.context.pageContext.web.serverRelativeUrl == "/")
      rootSiteURL = this.props.context.pageContext.web.absoluteUrl;
    else
      rootSiteURL = this.props.siteUrl;
    sp.setup({
      spfxContext: this.props.context as any
    });
    spweb = Web(rootSiteURL + '/' + this.state.inclusionpath + "/" + siteconfig.sitename);
  }

  public componentDidMount() {
    this.setState({
      isShow: true,
    });
    //check if TOT is already enabled
    this.checkTOTIsEnabled();

    //setting href on 'skip to main content' to navigate to main content
    const skipToMainContent = document.getElementById("sp-skipToContent");
    skipToMainContent.setAttribute("href", "#skipToMainContent");
    //adding span as a first element to the body to fix NVDA accessibility issue
    const spanElement = document.createElement("span");
    spanElement.innerHTML = "&nbsp;";
    spanElement.setAttribute("style", "height: 0px;display: block;");
    document.body.insertBefore(spanElement, skipToMainContent);

    //Get current user details and set state
    this.props.context.spHttpClient
      .get("/_api/SP.UserProfiles.PeopleManager/GetMyProperties", SPHttpClient.configurations.v1)
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          this.setState({ loggedinUserName: datauser.DisplayName });
          this.setState({ loggedinUserEmail: datauser.Email });
          //Create site and lists when app is installed 
          this.createSiteAndLists().then(() => {
            //Check current user's role and set UI components
            this.checkUserRole(datauser.Email);
            //Get list of Members from member List
            this.getMembersListData();
            //Get pending events from event track details list
            this.getPendingEventsData();
            //get appTitle from Config List
            this.getConfigData();
          });
          let props: any = {};
          datauser.UserProfileProperties.forEach((prop: any) => {
            props[prop.Key] = prop.Value;
          });
          datauser.userProperties = props;
          this.setState({ firstName: datauser.userProperties.FirstName });
        });
      }).catch((error: any) => {
        alert(stringsConstants.CMPErrorMessage + "while retrieving user details. Below is the " + JSON.stringify(error));
        console.error("CMP_CLBHome_componentDidMount_FailedToGetUserDetails \n", JSON.stringify(error));
      });

    //Get current Teams theme to pass into fluent provider
    app.initialize();
    app.getContext().then((ctx: any) => {
      const theme = ctx.theme || stringsConstants.themeDefaultMode;
      this.updateCMPTheme(theme);
    });
    app.registerOnThemeChangeHandler((theme: string) => {
      this.updateCMPTheme(theme);
    });

  }

  //Get Current theme when Teams theme is switched
  protected updateCMPTheme = (theme: string) => {
    switch (theme.toLocaleLowerCase()) {
      case stringsConstants.themeDefaultMode:
        this.setState({ currentTeamsTheme: teamsLightTheme, currentThemeName: stringsConstants.themeDefaultMode });
        break;
      case stringsConstants.themeDarkMode:
        this.setState({ currentTeamsTheme: teamsDarkTheme, currentThemeName: stringsConstants.themeDarkMode });
        break;
      case stringsConstants.themeContrastMode:
        this.setState({ currentTeamsTheme: teamsHighContrastTheme, currentThemeName: stringsConstants.themeContrastMode });
        break;
    }
  };

  //verify tot is enabled
  private async checkTOTIsEnabled() {
    try {
      const listStructure: any = siteconfig.totLists;
      let listsPresent = [];
      let allTOTLists = [];
      let fieldsMisMatchCount = 0;
      await listStructure.forEach(async (element: any) => {
        const spListTitle: string = element["listName"];
        const fieldsToCreate: string[] = element["fields"];
        allTOTLists.push(spListTitle);
        await spweb.lists.getByTitle(spListTitle).get().then(async (list: any) => {
          if (list != undefined) {
            listsPresent.push(spListTitle);
            //validate for fields
            let totalFieldsToCreate = await this.checkFieldExists(spListTitle, fieldsToCreate);
            if (totalFieldsToCreate.length != 0) {
              fieldsMisMatchCount++;
            }
          }
        })
          .catch(() => {
            this.setState({ isTOTEnabled: false });
          });
        //if mismatch found across lists
        if (listsPresent.length == 0 || (listsPresent.length != allTOTLists.length) || fieldsMisMatchCount > 0) {
          this.setState({ isTOTEnabled: false });
        }
        else {
          this.setState({ isTOTEnabled: true });
        }
      });
    }
    catch (error) {
      alert(stringsConstants.CMPErrorMessage + "while verifying tournament of Teams is enabled. Below is the " + JSON.stringify(error));
      console.error("CMP_CLBHome_checkTOTIsEnabled_FailedToVerifyListsFields \n", JSON.stringify(error));

    }

  }

  //validate if the list column already exists
  private async checkFieldExists(spListTitle: string, fieldsToCreate: string[]) {
    let totalFieldsToCreate = [];
    try {
      const filterFields = await spweb.lists.getByTitle(spListTitle).fields
        .filter("Hidden eq false and ReadOnlyField eq false")
        .get();
      for (let i = 0; i < fieldsToCreate.length; i++) {
        // compare fields 
        const parser = new DOMParser();
        const xml = parser.parseFromString(fieldsToCreate[i], 'text/xml');
        let fieldNameToCheck = xml.querySelector('Field').getAttribute('DisplayName');
        let fieldExists = filterFields.filter((e: any) => e.Title == fieldNameToCheck);
        if (fieldExists.length == 0) {
          totalFieldsToCreate.push(fieldsToCreate[i]);
        }
      }
      return totalFieldsToCreate;
    }
    catch (error) {
      alert(stringsConstants.CMPErrorMessage + "while checking required fields exists. Below are the details: \n" + JSON.stringify(error));
      console.error("CMP_clbhome_checkFieldExists \n", error);
    }
  }

  //create lists and Columns
  private async createNewList(item: any) {
    try {
      //Create List
      const listReponse = await commonService.createSPlist(item.displayName);
      console.log(stringsConstants.CMPLog + "List created: " + "'" + item.displayName + "'");
      //Create Columns
      const fieldsResponse = await commonService.createListFields(item.displayName, item.columns);
      console.log(stringsConstants.CMPLog + item.displayName + " List columns created");
      console.log("Responses", listReponse, fieldsResponse);
      if (listReponse && fieldsResponse === "Success") {
        setTimeout(() => {
          if (item.displayName === stringsConstants.MemberList) {
            //Add current user as a manager to the Champion Management Platform
            this.props.context.spHttpClient
              .get(
                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                SPHttpClient.configurations.v1
              )
              .then((userProperties: SPHttpClientResponse) => {
                userProperties.json().then(async (adminUser: any) => {
                  if (adminUser) {
                    this.props.context.spHttpClient
                      .get(
                        "/" + this.state.inclusionpath + "/" + this.state.sitename + `/_api/web/lists/GetByTitle('Member List')/Items?$filter= Title eq '${adminUser.Email}'`,
                        SPHttpClient.configurations.v1
                      )
                      .then((response: SPHttpClientResponse) => {
                        response.json().then((datada: any) => {
                          if (!datada.error) {
                            let val = datada.value;
                            if (val.length === 0 && item.displayName === stringsConstants.MemberList) {
                              const listDefinition: any = {
                                Title: adminUser.Email,
                                FirstName: adminUser.DisplayName.split(
                                  " "
                                )[0],
                                LastName: adminUser.DisplayName.split(
                                  " "
                                )[1],
                                Role: "Manager",
                                Status: "Approved",
                                Group: "IT Pro",
                                FocusArea: ['All'],
                              };
                              const spHttpClientOptions: ISPHttpClientOptions = {
                                body: JSON.stringify(listDefinition)
                              };
                              const url: string =
                                "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/items";
                              this.props.context.spHttpClient
                                .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
                                .then((newUserResponse: SPHttpClientResponse) => {
                                  this.setState({ isShow: false });
                                  if (newUserResponse.status === 201) {
                                    alert(
                                      ` ${this.state.loggedinUserName} has been added as a manager to the Champion Management Platform, please refresh the app to complete the setup.`
                                    );
                                  } else {
                                    alert("Response status " + newUserResponse.status + " - " + newUserResponse.statusText);
                                  }
                                }
                                );
                            }
                          }
                        });
                      });
                  }

                });
              }
              );
          }
        }, 6000);
        setTimeout(() => {
          if (item.displayName === stringsConstants.EventsList) {
            siteconfig.eventsMasterData.forEach(
              (eventData) => {
                let eventDataList: any = {
                  Title: eventData.Title,
                  Points: eventData.Points,
                  Description: eventData.Description,
                  IsActive: true,
                };
                const spHttpClientOptions: ISPHttpClientOptions = {
                  body: JSON.stringify(eventDataList),
                };

                const url: string = "/" + this.state.inclusionpath + "/" + this.state.sitename +
                  "/_api/web/lists/GetByTitle('Events List')/items";
                this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
              }
            );
          }
          //Creating items in Config List
          if (item.displayName === stringsConstants.ConfigList) {
            siteconfig.configMasterData.forEach(
              (configData) => {
                let configList: any = {
                  Title: configData.Title,
                  Value: configData.Value
                };
                const spHttpClientOptions: ISPHttpClientOptions = {
                  body: JSON.stringify(configList),
                };

                const url: string =
                  "/" +
                  this.state.inclusionpath +
                  "/" +
                  this.state.sitename +
                  "/_api/web/lists/GetByTitle('" + stringsConstants.ConfigList + "')/items";
                this.props.context.spHttpClient
                  .post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
              }
            );
          }
        }, 5000);
      }
      else {
        throw new Error(listReponse + "\n" + fieldsResponse);
      }
    }
    catch (error) {
      alert(stringsConstants.CMPErrorMessage + "while creating new list. Below are the details: \n" + JSON.stringify(error));
      console.error("CMP_CLBHome_createNewList_FailedtoCreateList \n", JSON.stringify(error));
    }
  }

  //When app is installed create new site collection and lists if not existing already
  private async createSiteAndLists() {
    console.log(stringsConstants.CMPLog + "Checking if site exists already.");
    //Set Variables
    let exSiteId: any;
    try {
      //Check if CMP site exists        
      await sp.site.exists(rootSiteURL + "/" + this.state.inclusionpath + "/" + this.state.sitename).then(async (response) => {
        if (response != undefined) {
          //If CMP site does not exist, create the site and lists
          if (!response) {
            flagCheckUserRole = false;
            console.log(stringsConstants.CMPLog + "Creating new site collection: '" + this.state.sitename + "'");
            //Create a new site collection
            const createSiteUrl: string = "/_api/SPSiteManager/create";
            const siteDefinition: any = {
              request: {
                Title: this.state.sitename,
                Url:
                  this.state.siteUrl.replace("https:/", "https://").replace("https:///", "https://") +
                  "/" +
                  this.state.inclusionpath +
                  "/" +
                  this.state.sitename,
                Lcid: 1033,
                ShareByEmailEnabled: true,
                Description: "Description",
                WebTemplate: "STS#3",
                SiteDesignId: "6142d2a0-63a5-4ba0-aede-d9fefca2c767",
                Owner: this.state.loggedinUserEmail,
              },
            };
            const spHttpsiteClientOptions: ISPHttpClientOptions = {
              body: JSON.stringify(siteDefinition),
            };

            //HTTP post request for creating a new site collection
            this.props.context.spHttpClient
              .post(
                createSiteUrl,
                SPHttpClient.configurations.v1,
                spHttpsiteClientOptions
              )
              .then(async (siteResponse: SPHttpClientResponse) => {
                //If site is succesfully created
                if (siteResponse.status === 200) {
                  console.log(stringsConstants.CMPLog + "Created new site collection: '" + this.state.sitename + "'");
                  //site is newly created and got 200 response, now create Digital Lib
                  await this.createDigitalBadgeLib();
                  //create CMP Logo Library to store the organization logo that allows users to customize it.
                  await this.getAppLogoImage();
                  siteResponse.json().then((siteData: any) => {
                    if (siteData.SiteId) {
                      exSiteId = siteData.SiteId;
                      this.setState({ siteId: siteData.SiteId });
                      console.log(stringsConstants.CMPLog + "Creating Lists in new site");
                      //Create 3 lists in the newly created site
                      if (exSiteId) {
                        let lists: any = [];
                        siteconfig.lists.forEach((spList) => {
                          let list = {
                            displayName: spList.listName,
                            columns: spList.fields
                          };
                          lists.push(list);
                        });
                        commonService = new CommonServices(this.props.context, this.props.siteUrl);
                        lists.forEach((spList: any) => {
                          this.createNewList(spList);
                        });
                      }
                    }
                  });
                }

              }).catch((error: any) => {
                alert(stringsConstants.CMPErrorMessage + "while creating new site. Below are the details: \n" + JSON.stringify(error));
                console.error("CMP_CLBHome_createSiteAndLists_FailedToCreateSite \n", JSON.stringify(error));
              });

          }//IF END
          //If CMP site already exists create only lists.  
          else {
            //site exists, check if Digital Badge lib exists, create if not present
            await this.createDigitalBadgeLib();
            //check if CMP Logo lib exists, create if not present
            await this.getAppLogoImage();
            //Check if Lists exists already
            this.props.context.spHttpClient
              .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/Items", SPHttpClient.configurations.v1)
              .then(async (responseMemberList: SPHttpClientResponse) => {
                if (responseMemberList.status === 404) {
                  //If lists do not exist create lists. Else no action is required
                  console.log(stringsConstants.CMPLog + "Site already existing but lists not found");
                  console.log(stringsConstants.CMPLog + "Getting site collection ID for creating lists");
                  flagCheckUserRole = false;
                  //Get Sitecollection ID for creating lists   
                  this.props.context.spHttpClient
                    .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/site/id", SPHttpClient.configurations.v1)
                    .then((responseuser: SPHttpClientResponse) => {
                      if (responseuser.status === 404) {
                        alert(stringsConstants.CMPErrorMessage + "while setting up the App. Please try refreshing or loading after some time.");
                        console.error("CMP_CLBHome_createSiteAndLists_FailedToGetSiteID \n");
                      }
                      else {
                        responseuser.json().then((datauser: any) => {
                          exSiteId = datauser.value;
                          if (exSiteId) {
                            console.log(stringsConstants.CMPLog + "Creating lists");
                            //Set up List Creation Information
                            let lists: any = [];
                            siteconfig.lists.forEach((spList) => {
                              let list = {
                                displayName: spList.listName,
                                columns: spList.fields
                              };
                              lists.push(list);
                            });
                            commonService = new CommonServices(this.props.context, this.props.siteUrl);
                            //Iterate and create all the lists
                            lists.forEach((spList: any) => {
                              this.createNewList(spList);
                            });
                          }
                        });
                      }
                    }).catch((error: any) => {
                      alert(stringsConstants.CMPErrorMessage + "while retrieving SiteID. Below are the details: \n" + JSON.stringify(error));
                      console.error("CMP_CLBHome_createSiteAndLists_FailedToGetSiteID \n", JSON.stringify(error));
                    });

                }
                else if (responseMemberList.status === 200) {
                  //If Member list exists, modify the Focus Area Choice column from Dropdown to Checkboxes during app upgrade 
                  const focusAreaColumn = await spweb.lists.getByTitle(stringsConstants.MemberList).fields.getByInternalNameOrTitle(stringsConstants.FocusAreaColumn)();
                  if (focusAreaColumn.TypeAsString == "Choice") {
                    await spweb.lists.getByTitle(stringsConstants.MemberList).fields.getByInternalNameOrTitle(stringsConstants.FocusAreaColumn).update({ TypeAsString: "MultiChoice" });
                    console.log("CMP_CLBHome_FocusArea column type is updated successsfully");
                  }
                  //Add Status and Notes field to Event Track Details List, if not present. 
                  const eventTrackDetailsColumns = await spweb.lists.getByTitle(stringsConstants.EventTrackDetailsList).fields.filter("Hidden eq false and ReadOnlyField eq false").select("Title").get();

                  //Check if Notes column exists, if not create it.
                  let notesFieldExists = eventTrackDetailsColumns.filter((e: any) => e.Title == stringsConstants.NotesColumn);
                  if (notesFieldExists.length == 0) {
                    await spweb.lists.getByTitle(stringsConstants.EventTrackDetailsList).fields.addMultilineText(stringsConstants.NotesColumn, 6, false, false, false, false).then(async () => {
                      await spweb.lists.getByTitle(stringsConstants.EventTrackDetailsList).defaultView.fields.add(stringsConstants.NotesColumn).then(() => {
                        console.log("CMP_CLBHome_Added Notes column to Event Track Details List");
                      });
                    }).catch((error: any) => {
                      console.error("CMP_CLBHome_Failed to add Notes Column to Event Track Details List \n", JSON.stringify(error));
                    });
                  }

                  //Check if Status column exists, if not create it.
                  let statusFieldExists = eventTrackDetailsColumns.filter((e: any) => e.Title == stringsConstants.StatusColumn);
                  if (statusFieldExists.length == 0) {
                    await spweb.lists.getByTitle(stringsConstants.EventTrackDetailsList).fields.addText(stringsConstants.StatusColumn, 255).then(async () => {
                      await spweb.lists.getByTitle(stringsConstants.EventTrackDetailsList).defaultView.fields.add(stringsConstants.StatusColumn).then(() => {
                        console.log("CMP_CLBHome_Added Status column to Event Track Details List");
                      });
                    }).catch((error: any) => {
                      console.error("CMP_CLBHome_Failed to add Status Column to Event Track Details List \n", JSON.stringify(error));
                    });
                  }

                  //Add Config List to the CMP site
                  let configList = {
                    displayName: stringsConstants.ConfigList,
                    columns: siteconfig.lists.find((item) => item.listName === stringsConstants.ConfigList).fields
                  };

                  //Get Sitecollection ID for creating Config list   
                  this.props.context.spHttpClient
                    .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/site/id", SPHttpClient.configurations.v1)
                    .then((responseSiteId: SPHttpClientResponse) => {
                      if (responseSiteId.status === 404) {
                        alert(stringsConstants.CMPErrorMessage + "while setting up the App. Please try refreshing or loading after some time.");
                        console.error("CMP_CLBHome_createSiteAndLists_FailedToGetSiteID \n");
                      }
                      else {
                        responseSiteId.json().then((datauser: any) => {
                          exSiteId = datauser.value;
                          console.log("exSiteId: " + exSiteId);
                          this.props.context.spHttpClient
                            .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Config List')/Items", SPHttpClient.configurations.v1)
                            .then(async (responseConfigList: SPHttpClientResponse) => {
                              if (responseConfigList.status === 404) {
                                commonService = new CommonServices(this.props.context, this.props.siteUrl);
                                this.createNewList(configList);
                                console.log("CMP_CLBHome_Config List was created");
                              }
                              else {
                                //creating the default entries in the config list, if not exists.
                                let configItems = await spweb.lists.getByTitle(stringsConstants.ConfigList).items.getAll();

                                siteconfig.configMasterData.forEach(async (configData) => {
                                  if (configItems.length === 0 ||
                                    configItems.find((item: any) => item.Title === configData.Title) === undefined) {
                                    let configListItem: any = {
                                      Title: configData.Title,
                                      Value: configData.Value
                                    };
                                    await spweb.lists.getByTitle(stringsConstants.ConfigList).items.add(configListItem).then().catch((error: any) => {
                                      console.error("CMP_CLBHome_Failed to add item to Config List \n", JSON.stringify(error));
                                    });
                                  }
                                })
                              }
                            });
                        });
                      }
                    });

                }
              }).catch((error: any) => {
                alert(stringsConstants.CMPErrorMessage + "while checking if MemberList exists. Below are the details: \n" + JSON.stringify(error));
                console.error("CMP_CLBHome_createSiteAndLists_FailedToCheckMemberListExists \n", JSON.stringify(error));
              });

          }//End of else part - CMP site already exists create only lists. 

        }//First IF END

      }).catch((error) => {
        alert(stringsConstants.CMPErrorMessage + "while checking if site exists. Below are the details: \n" + JSON.stringify(error));
        console.error("CMP_CLBHome_createSiteAndLists_FailedtoCheckIfSiteExists \n", JSON.stringify(error));
      });
    }
    catch (error) {
      console.error("CMP_CLBHome_createSiteAndLists \n", error);
      alert(stringsConstants.CMPErrorMessage + "while creating site and lists. Below are the details: \n" + error);
    }
  }

  //Create CMP Logo library to store organization logo
  private async getAppLogoImage(): Promise<any> {
    try {
      let logoImageURL: string;
      const spListTitle: string = stringsConstants.CMPLogoLibrary;
      //check if CMP Logo lib exists, create if doesn't exists and upload default logo 
      await spweb.lists.getByTitle(spListTitle).get().then(async () => {
        var logoImageArray: any[] = [];
        logoImageArray = await spweb.lists.getByTitle(spListTitle).items.select("File/Name").expand("File").get();
        if (logoImageArray.length > 0) {
          for (let i = 0; i < logoImageArray.length; i++) {
            if (logoImageArray[i].File.Name.toLowerCase() == stringsConstants.AppLogoLowerCase) {
              logoImageURL = rootSiteURL + "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/" + spListTitle + "/" + logoImageArray[i].File.Name;
              this.setState({
                appLogoURL: logoImageURL,
              });
            }
            else {
              this.setState({
                appLogoURL: require("../assets/CMPImages/" + stringsConstants.AppLogo),
              });
            }
          }
        } else {
          this.setState({
            appLogoURL: require("../assets/CMPImages/" + stringsConstants.AppLogo),
          });
        }
      }).catch(async () => {
        //create library in SharePoint site to store the organization logo image
        await spweb.lists.add(spListTitle, "", 101, true).then(async () => {
          fetch(require("../assets/CMPImages/" + stringsConstants.AppLogo)).then(res => res.blob()).then((blob) => {
            spweb.getFolderByServerRelativeUrl("/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/" + spListTitle).files.add(stringsConstants.AppLogo, blob, true);
            this.setState({
              appLogoURL: require("../assets/CMPImages/" + stringsConstants.AppLogo),
            });
          });
        });
      });  //catch end         
    }
    catch (error) {
      console.error("CMP_CLBHome_getAppLogoImage \n", error);
    }
  }
  //create digital lib to store the badges
  private async createDigitalBadgeLib(): Promise<any> {
    try {
      const listStructure: any = siteconfig.libraries;
      for (let i = 0; i < listStructure.length; i++) {
        const spListTitle: string = listStructure[i]["listName"];
        const spListTemplate = listStructure[i]["listTemplate"];
        //check if digital assests lib exists, create if doesn't exists and upload default badge 
        await spweb.lists.getByTitle(spListTitle).get().then(async () => {
          //Check if "MinimumPoints" field exists, if not create it
          spweb.lists.getByTitle(spListTitle).fields.getByInternalNameOrTitle(stringsConstants.MinimumPointsColumn).get()
            .then(() => {

            }).catch(async () => {
              //field doesn't exists, hence create it
              await spweb.lists.getByTitle(spListTitle).fields.addNumber(stringsConstants.MinimumPointsColumn, 0).then(async () => {
                await spweb.lists.getByTitle(spListTitle).defaultView.fields.add(stringsConstants.MinimumPointsColumn).then(() => {
                  console.log("CMP_CLBHome_Added Minimum Points column to Digital Badge Library");
                });
              });
            });
          //Add "Title" to default view
          const defaultXML = await spweb.lists.getByTitle(spListTitle).defaultView.fields.getSchemaXml();
          let titleFieldIndex = defaultXML.indexOf("Title");
          if (titleFieldIndex == -1) {
            await spweb.lists.getByTitle(spListTitle).defaultView.fields.add("Title");
          }
        }).
          catch(async () => {
            //create lib
            await spweb.lists.add(spListTitle, "", spListTemplate, true).then(async () => {
              //Create "MinimumPoints" field
              await spweb.lists.getByTitle(spListTitle).fields.addNumber(stringsConstants.MinimumPointsColumn, 0).then(async () => {
                await spweb.lists.getByTitle(spListTitle).defaultView.fields.add(stringsConstants.MinimumPointsColumn).then(() => {
                  console.log("CMP_CLBHome_Added Minimum Points column to Digital Badge Library");
                });
              });
              fetch(require('../assets/images/CMPBadge.png')).then(res => res.blob()).then((blob) => {
                spweb.getFolderByServerRelativeUrl("/" + this.state.inclusionpath + "/" + this.state.sitename + "/" + spListTitle).files.add("digitalbadge.png", blob, true)
                  .then((res: any) => {
                    res.file.getItem().then((item: any) => {
                      item.update({
                        Title: "Teamwork Champion"
                      });
                    });
                  });
              });
              const defaultXML = await spweb.lists.getByTitle(spListTitle).defaultView.fields.getSchemaXml();
              let titleFieldIndex = defaultXML.indexOf("Title");
              if (titleFieldIndex == -1) {
                await spweb.lists.getByTitle(spListTitle).defaultView.fields.add("Title");
              }
            });
          });  //catch end    
      }
    }
    catch (error) {
      console.error("CMP_CLBHome_createDigitalBadgeLib \n", error);
    }
  }

  //Check current users's role from "Member List" and set the UI components accordingly
  private async checkUserRole(userEmail: string) {
    try {
      if (flagCheckUserRole) {

        console.log(stringsConstants.CMPLog + "Checking user role and setting the UI components");
        this.props.context.spHttpClient
          .get(
            "/" +
            this.state.inclusionpath +
            "/" +
            this.state.sitename +
            "/_api/web/lists/GetByTitle('Member List')/Items?$filter=Title eq '" + userEmail.toLowerCase() + "'",
            SPHttpClient.configurations.v1
          )
          .then((response: SPHttpClientResponse) => {
            if (response.status === 200) {
              this.setState({
                isShow: false,
              });
            }
            response.json().then((datada: any) => {
              if (!datada.error) {
                let dataexists: any = datada.value.find(
                  (x: any) =>
                    x.Title.toLowerCase() === this.state.loggedinUserEmail.toLowerCase()
                );
                if (dataexists) {
                  if (dataexists.Status === "Approved") {
                    if (dataexists.Role === "Manager") {
                      this.setState({ clB: true });
                    } else if (dataexists.Role === "Champion") {
                      this.setState({ cV: true });
                    }
                  } else if (
                    dataexists.Role === "Employee" ||
                    dataexists.Role === "Champion" ||
                    dataexists.Role === "Manager"
                  ) {
                    this.setState({ eV: true });
                  }
                }
                else
                  this.setState({ eV: true });
              }
            });
          }).catch((error: any) => {
            alert(stringsConstants.CMPErrorMessage + "while retrieving user role. Below are the details: \n" + JSON.stringify(error));
            console.error("CMP_CLBHome_checkUserRole_FailedtoGetUserRole \n", JSON.stringify(error));
          });
      }
    }
    catch (error) {
      console.error("CMP_CLBHome_checkUserRole \n", error);
      alert(stringsConstants.CMPErrorMessage + " while retrieving user role. Below are the details: \n" + error);
    }
  }

  //Get the list of Members from member List
  private getMembersListData(): void {
    this.props.context.spHttpClient
      .get(
        "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/_api/web/lists/GetByTitle('" + stringsConstants.MemberList + "')/Items?$top=1&$filter= Status eq 'Pending'",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          response.json().then((responseJSON: any) => {
            this.setState({
              isChampionApprovalPending: responseJSON.value.length > 0
            });
          });
        }
      });
  }

  //get Config List data
  private async getConfigData(): Promise<void> {
    try {
      let filterQuery = "Title eq '" + stringsConstants.AppTitle + "'";
      const configListData = await spweb.lists.getByTitle(stringsConstants.ConfigList).items.filter(filterQuery).getAll();

      if (configListData.length > 0) {
        this.setState({ appTitle: configListData[0].Value });
      }
      else {
        this.setState({ appTitle: siteconfig.appTitle });
      }
    } catch (error) {
      console.error("CMP_clbhome_getConfigData \n", error);
    }
  }
  //Check for any pending approvals in the event track details list
  private getPendingEventsData(): void {
    this.props.context.spHttpClient
      .get(
        "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/_api/web/lists/GetByTitle('" + stringsConstants.EventTrackDetailsList + "')/Items?$top=1&$filter= Status eq 'Pending'",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          response.json().then((responseJSON: any) => {
            this.setState({
              isEventApprovalPending: responseJSON.value.length > 0
            });
          });
        }
      });
  }

  //callback function
  private callBackFunction(): void {
    try {
      //Get the list of Members from member List
      this.getMembersListData();

      //Get pending events from event track details list
      this.getPendingEventsData();

      this.setState({
        cB: false,
        ChampionsList: false,
        addMember: false,
        manageApprovals: false
      });
    } catch (error) {
      console.log(error);
    }
  }

  //Get CMP Page Banner Class when Theme is switched to Dark or Contrast Mode.
  public getCMPSubHeaderBannerClass() {
    if (this.state.currentThemeName === stringsConstants.themeDarkMode) {
      return styles.cmpHeaderDark;
    }
    else if (this.state.currentThemeName === stringsConstants.themeContrastMode) {
      return styles.cmpHeaderContrast;
    }
    else
      return "";
  }

  public render(): React.ReactElement<IClbHomeProps> {
    const isDarkOrContrastTheme = this.state.currentThemeName === stringsConstants.themeDarkMode || this.state.currentThemeName === stringsConstants.themeContrastMode;
    return (
      <FluentProvider className={styles.clbHome} theme={this.state.currentTeamsTheme}>
        {this.state.isShow && <div className={styles.load}></div>}
        < div className={styles.container} >
          <div>
            <Header
              logoImageURL={this.state.appLogoURL}
              appTitle={this.state.appTitle}
              showSearch={this.state.cB}
              clickcallback={() =>
                this.setState({
                  cB: false,
                  ChampionsList: false,
                  addMember: false,
                  dB: false,
                  manageApprovals: false,
                  enableTOT: false,
                  championReport: false

                })
              }
              currentThemeName={this.state.currentThemeName}
            />
          </div>
          <main aria-label={this.state.appTitle} id="skipToMainContent">
            {!this.state.cB &&
              !this.state.ChampionsList &&
              !this.state.addMember && !this.state.manageApprovals &&
              !this.state.dB && !this.state.enableTOT && !this.state.championReport && (
                <div>
                  <div className={`${styles.imgheader} ${this.getCMPSubHeaderBannerClass()}`.trim()}>
                    <span className={styles.cmpPageHeading}>{LocaleStrings.WelcomeLabel} {this.state.firstName}!</span>
                  </div>
                  <div className={styles.grid}>
                    <h2><div className={`${styles.quickguide}${isDarkOrContrastTheme ? " " + styles.quickguideDarkContrast : ""}`}>
                      {LocaleStrings.GetStartedLabel}</div></h2>
                    <Row xl={4} lg={4} md={4} sm={3} xs={2} className="mt-4">
                      <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                        <Media
                          className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                          onClick={() => this.setState({ cB: !this.state.cB })}
                          onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.setState({ cB: !this.state.cB }) }}
                        >
                          <div className={styles.mb}>
                            <img
                              src={require("../assets/CMPImages/ChampionLeaderBoard.svg")}
                              alt={LocaleStrings.ChampionLeaderBoardLabel}
                              className={styles.dashboardimgs}
                              title={LocaleStrings.ChampionLeaderBoardLabel}
                              aria-label={LocaleStrings.ChampionLeaderBoardLabel}
                              role="button"
                              tabIndex={0}
                            />
                            <div className={styles.center} aria-hidden="true" title={LocaleStrings.ChampionLeaderBoardLabel}>
                              {LocaleStrings.ChampionLeaderBoardLabel}
                            </div>
                          </div>
                        </Media>
                      </Col>
                      {(this.state.cV || this.state.clB) && (
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media
                            className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                            onKeyDown={(evt: any) => {
                              if (evt.key === stringsConstants.stringEnter) this.setState({
                                addMember: !this.state.addMember,
                              })
                            }}
                            onClick={() =>
                              this.setState({
                                addMember: !this.state.addMember,
                              })
                            }
                          >
                            <div className={styles.mb} >
                              <img
                                src={require("../assets/CMPImages/AddMembers.svg")}
                                alt={LocaleStrings.AddMembersToolTip}
                                className={styles.dashboardimgs}
                                aria-label={(this.state.clB && !this.state.cV) ? LocaleStrings.AddMemberLabel : LocaleStrings.NominateMemberLabel}
                                role="button"
                                tabIndex={0}
                                title={LocaleStrings.AddMembersToolTip}
                              />
                              <div className={styles.center} aria-hidden="true" title={(this.state.clB && !this.state.cV) ? LocaleStrings.AddMemberLabel : LocaleStrings.NominateMemberLabel}>
                                {(this.state.clB && !this.state.cV) ? LocaleStrings.AddMemberLabel : LocaleStrings.NominateMemberLabel}
                              </div>
                            </div>
                          </Media>
                        </Col>
                      )}
                      {(this.state.cV || this.state.clB) && (
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media
                            className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                            onClick={() => this.setState({ dB: !this.state.dB })}
                            onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.setState({ dB: !this.state.dB }) }}
                          >
                            <div className={styles.mb} >
                              <img
                                src={require("../assets/CMPImages/DigitalBadge.svg")}
                                alt={LocaleStrings.DigitalMembersToolTip}
                                className={styles.dashboardimgs}
                                aria-label={LocaleStrings.DigitalBadgeLabel}
                                role="button"
                                tabIndex={0}
                                title={LocaleStrings.DigitalMembersToolTip}
                              />
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.DigitalBadgeLabel}>{LocaleStrings.DigitalBadgeLabel}</div>
                            </div>
                          </Media>
                        </Col>
                      )}
                      {this.state.isTOTEnabled && (
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <div>
                            <Media
                              className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                              onClick={() => this.setState({ enableTOT: !this.state.enableTOT })}
                              onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.setState({ enableTOT: !this.state.enableTOT }) }}
                            >
                              <div className={styles.mb} >
                                <img
                                  src={require("../assets/CMPImages/TournamentOfTeams.svg")}
                                  alt={LocaleStrings.TOTLabel}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.TOTLabel}
                                  aria-label={LocaleStrings.TOTLabel}
                                  role="button"
                                  tabIndex={0}
                                />
                                {this.state.isTOTEnabled && (<div className={`${styles.center} ${styles.totLabel}`} aria-hidden="true" title={LocaleStrings.TOTLabel}>
                                  {LocaleStrings.TOTLabel}</div>)}
                              </div>
                            </Media>
                          </div>
                        </Col>)}
                    </Row>

                    {this.state.clB && !this.state.cV && (
                      <h2><div className={`${styles.admintools}${isDarkOrContrastTheme ? " " + styles.admintoolsDarkContrast : ""}`}>
                        {LocaleStrings.AdminToolsLabel}</div></h2>)}

                    {this.state.clB && !this.state.cV && (
                      <Row xl={4} lg={4} md={4} sm={3} xs={2} className="mt-4">
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                            onKeyDown={(evt: any) => {
                              if (evt.key === stringsConstants.stringEnter) this.setState({
                                manageApprovals: !this.state.manageApprovals,
                              })
                            }}

                            onClick={() =>
                              this.setState({
                                manageApprovals: !this.state.manageApprovals,
                              })
                            }
                          >
                            <div className={styles.mb} >
                              <img
                                src={(this.state.isChampionApprovalPending || this.state.isEventApprovalPending) ?
                                  require("../assets/CMPImages/ManagePendingApprovals.svg") :
                                  require("../assets/CMPImages/ManageApprovals.svg")}
                                alt={LocaleStrings.AdminTasksLabel}
                                className={styles.dashboardimgs}
                                title={LocaleStrings.AdminTasksLabel}
                                aria-label={LocaleStrings.AdminTasksLabel}
                                role="button"
                                tabIndex={0}
                              />
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.AdminTasksLabel}>
                                {LocaleStrings.AdminTasksLabel}
                              </div>
                            </div>
                          </Media>
                        </Col>
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}>
                            <div className={styles.mb}>
                              <a
                                href={`/${this.state.inclusionpath}/${this.state.sitename}/Lists/Member%20List/AllItems.aspx`}
                                target="_blank"

                              >
                                <img
                                  src={require("../assets/CMPImages/ChampionList.svg")}
                                  alt={LocaleStrings.ChampionsListToolTip}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.ChampionsListToolTip}
                                  aria-label={LocaleStrings.ChampionListLabel}
                                  role="button"
                                />
                              </a>
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.ChampionListLabel}>{LocaleStrings.ChampionListLabel}</div>
                            </div>
                          </Media>
                        </Col>
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}>
                            <div className={styles.mb} >
                              <a
                                href={`/${this.state.inclusionpath}/${this.state.sitename}/Lists/Events%20List/AllItems.aspx`}
                                target="_blank"
                              >
                                <img
                                  src={require("../assets/CMPImages/EventsList.svg")}
                                  alt={LocaleStrings.EventsListToolTip}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.EventsListToolTip}
                                  aria-label={LocaleStrings.EventsListLabel}
                                  role="button"
                                />
                              </a>
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.EventsListLabel}>{LocaleStrings.EventsListLabel}</div>
                            </div>
                          </Media>
                        </Col>
                        {!this.state.isTOTEnabled && (
                          <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                            <Media
                              className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                              onClick={() => this.setState({ enableTOT: !this.state.enableTOT })}
                              onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.setState({ enableTOT: !this.state.enableTOT }) }}

                            >
                              <div className={styles.mb} >
                                <img
                                  src={require("../assets/CMPImages/EnableTOT.svg")}
                                  alt={LocaleStrings.EnableTOTToolTip}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.EnableTOTToolTip}
                                  aria-label={LocaleStrings.EnableTOTLabel}
                                  role="button"
                                  tabIndex={0}
                                />
                                {!this.state.isTOTEnabled && (<div className={`${styles.center} ${styles.enableTournamentLabel}`} title={LocaleStrings.EnableTOTLabel} aria-hidden="true">
                                  {LocaleStrings.EnableTOTLabel}</div>)}
                              </div>
                            </Media>
                          </Col>)}
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}>
                            <div className={styles.mb} >
                              <a
                                href={`/${this.state.inclusionpath}/${this.state.sitename}/Digital%20Badge%20Assets/Forms/AllItems.aspx`}
                                target="_blank"
                              >
                                <img
                                  src={require("../assets/CMPImages/ManageDigitalBadges.svg")}
                                  alt={LocaleStrings.ManageDigitalBadgesToolTip}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.ManageDigitalBadgesToolTip}
                                  aria-label={LocaleStrings.ManageDigitalBadgesLabel}
                                  role="button"
                                />
                              </a>
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.ManageDigitalBadgesLabel}>
                                {LocaleStrings.ManageDigitalBadgesLabel}
                              </div>
                            </div>
                          </Media>
                        </Col>
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}>
                            <div className={styles.mb} >
                              <a
                                href={`/${this.state.inclusionpath}/${this.state.sitename}/CMP%20Logo/Forms/AllItems.aspx`}
                                target="_blank"

                              >
                                <img
                                  src={require("../assets/CMPImages/ManageAppLogo.svg")}
                                  alt={LocaleStrings.ManageAppLogoToolTip}
                                  className={styles.dashboardimgs}
                                  title={LocaleStrings.ManageAppLogoToolTip}
                                  role="button"
                                  aria-label={LocaleStrings.ManageAppLogoLabel}
                                />
                              </a>
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.ManageAppLogoLabel}>
                                {LocaleStrings.ManageAppLogoLabel}
                              </div>
                            </div>
                          </Media>
                        </Col>
                        <Col xl={3} lg={3} md={3} sm={4} xs={6} className={styles.imageLayout}>
                          <Media
                            className={`${styles.cursor}${isDarkOrContrastTheme ? " " + styles.cursorDarkContrast : ""}`}
                            onClick={() => this.setState({ championReport: !this.state.championReport })}
                            onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.setState({ championReport: !this.state.championReport }) }}
                          >
                            <div className={styles.mb} >
                              <img
                                src={require("../assets/TOTImages/TournamentsReport.svg")}
                                alt={LocaleStrings.ChampionReportToolTip}
                                className={styles.dashboardimgs}
                                title={LocaleStrings.ChampionReportToolTip}
                                aria-label={LocaleStrings.ChampionReportLabel}
                                role="button"
                                tabIndex={0}
                              />
                              <div className={styles.center} aria-hidden="true" title={LocaleStrings.ChampionReportLabel}>{LocaleStrings.ChampionReportLabel}</div>
                            </div>
                          </Media>
                        </Col>
                      </Row>
                    )}
                  </div>
                </div>
              )
            }
            {
              this.state.cB && this.state.clB && (
                <ChampionLeaderBoard
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  onClickCancel={() => this.setState({ clB: true, cB: false })}
                  loggedinUserEmail={this.state.loggedinUserEmail}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.cB && this.state.cV && (
                <ChampionLeaderBoard
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  onClickCancel={() => this.setState({ clB: false, cB: false })}
                  loggedinUserEmail={this.state.loggedinUserEmail}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.enableTOT && (
                <TOTLandingPage
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  isTOTEnabled={this.state.isTOTEnabled}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.addMember && (
                <ClbAddMember
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  isAdmin={this.state.clB && !this.state.cV}
                  onClickBack={() => { this.setState({ addMember: false }); }}
                  onHomeCallBack={this.callBackFunction}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.manageApprovals && (
                <ManageApprovals
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  onClickBack={this.callBackFunction}
                  updateAppTitle={this.setState.bind(this)}
                  isPendingChampionApproval={this.state.isChampionApprovalPending}
                  isPendingEventApproval={this.state.isEventApprovalPending}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.cB && this.state.eV && (
                <EmployeeView
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  onClickCancel={() => this.setState({ eV: true, cB: false })}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.dB && (
                <DigitalBadge
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  clickcallback={() => this.setState({ dB: false })}
                  currentThemeName={this.state.currentThemeName}
                />
              )
            }
            {
              this.state.championReport && (

                <ChampionReport
                  siteUrl={this.props.siteUrl}
                  context={this.props.context}
                  appTitle={this.state.appTitle}
                  loggedinUserEmail={this.state.loggedinUserEmail}
                  onClickCancel={() => {
                    this.setState({ championReport: false });
                  }}
                />
              )}
          </main>
        </div>
      </FluentProvider>
    );
  }
}
