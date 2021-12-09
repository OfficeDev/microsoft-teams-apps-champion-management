import {
  ISPHttpClientOptions, MSGraphClient, SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import "@pnp/sp/lists";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/views";
import { Web } from "@pnp/sp/webs";
import { initializeIcons } from "@uifabric/icons";
import "bootstrap/dist/css/bootstrap.min.css";
import { ThemeStyle } from "msteams-ui-styles-core";
import * as React from "react";
import Col from "react-bootstrap/Col";
import Media from "react-bootstrap/Media";
import Row from "react-bootstrap/Row";
import siteconfig from "../config/siteconfig.json";
import styles from "../scss/CMPHome.module.scss";
import ApproveChampion from "./ApproveChampion";
import ChampionLeaderBoard from "./ChampionLeaderBoard";
import ClbAddMember from "./ClbAddMember";
import ClbChampionsList from "./ClbChampionsList";
import DigitalBadge from "./DigitalBadge";
import EmployeeView from "./EmployeeView";
import Header from "./Header";
import { IClbHomeProps } from "./IClbHomeProps";
import TOTLandingPage from "./TOTLandingPage";

initializeIcons();

export interface IClbHomeState {
  cB: boolean;
  clB: boolean;
  addMember: boolean;
  approveMember: boolean;
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
  isUserAdded: boolean;
  userStatus: string;
  firstName: string;
}

//Global Variables
const cmpLog: string = "CMP Logs: ";
let flagCheckUserRole: boolean = true;
const errorMessage: string = "An unexpected error occured ";
let rootSiteURL: string;
let spweb: any;

export default class ClbHome extends React.Component<
  IClbHomeProps,
  IClbHomeState
> {
  constructor(_props: any) {
    super(_props);
    this.state = {
      siteUrl: this.props.siteUrl,
      cB: false,
      addMember: false,
      approveMember: false,
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
      isUserAdded: false,
      userStatus: "",
      firstName: ""
    };
    this.checkUserRole = this.checkUserRole.bind(this);

    //Set the context for PNP
    sp.setup({
      spfxContext: this.props.context
    });

    //set context
    if (this.props.context.pageContext.web.serverRelativeUrl == "/")
      rootSiteURL = this.props.context.pageContext.web.absoluteUrl;
    else
      rootSiteURL = this.props.siteUrl;

    spweb = Web(rootSiteURL + '/' + this.state.inclusionpath + "/" + siteconfig.sitename);
  }

  public componentDidMount() {

    this.setState({
      isShow: true,
    });
    //check if TOT is already enabled
    this.checkTOTIsEnabled();
    //Get current user details and set state
    this.props.context.spHttpClient
      .get(

        "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          this.setState({ loggedinUserName: datauser.DisplayName });
          this.setState({ loggedinUserEmail: datauser.Email });

          //Create site and lists when app is installed 
          this.createSiteAndLists().then(() => {
            //Check current user's role and set UI components
            this.checkUserRole(datauser.Email);
          });
          var props = {};
          datauser.UserProfileProperties.forEach((prop) => {
            props[prop.Key] = prop.Value;
          });
          datauser.userProperties = props;
          this.setState({ firstName: datauser.userProperties.FirstName });
        });
      }).catch((error) => {
        alert(errorMessage + "while retrieving user details. Below is the " + JSON.stringify(error));
        console.error("CMP_CLBHome_componentDidMount_FailedToGetUserDetails \n", JSON.stringify(error));
      });

  }
  //verify tot is enabled
  private async checkTOTIsEnabled() {
    try {
      const listStructure: any = siteconfig.totLists;
      let listsPresent = [];
      let allTOTLists = [];
      let fieldsMisMatchCount = 0;
      await listStructure.forEach(async (element) => {
        const spListTitle: string = element["listName"];
        const fieldsToCreate: string[] = element["fields"];
        allTOTLists.push(spListTitle);
        await spweb.lists.getByTitle(spListTitle).get().then(async (list) => {
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
      alert(errorMessage + "while verifying tournament of Teams is enabled. Below is the " + JSON.stringify(error));
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
        let fieldExists = filterFields.filter(e => e.Title == fieldNameToCheck);
        if (fieldExists.length == 0) {
          totalFieldsToCreate.push(fieldsToCreate[i]);
        }
      }
      return totalFieldsToCreate;
    }
    catch (error) {
      alert(errorMessage + "while checking required fields exists. Below are the details: \n" + JSON.stringify(error));
      console.error("CMP_clbhome_checkFieldExists \n", error);
    }
  }

  //create lists 
  private createNewList(siteId: any, item: any) {
    this.props.context.msGraphClientFactory
      .getClient()
      .then(async (client: MSGraphClient) => {
        client
          .api("sites/" + siteId + "/lists")
          .version("v1.0")
          .header("Content-Type", "application/json")
          .responseType("json")
          .post(item, (errClbHome, _res, rawresponse) => {
            if (!errClbHome) {
              if (rawresponse.status === 201) {
                console.log(cmpLog + "List created: " + "'" + item.displayName + "'");
                setTimeout(() => {
                  this.props.context.spHttpClient
                    .get(

                      "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                      SPHttpClient.configurations.v1
                    )
                    .then(
                      (
                        userProperties: SPHttpClientResponse
                      ) => {
                        userProperties
                          .json()
                          .then(async (adminUser: any) => {

                            let goAndCreateMember: boolean;
                            goAndCreateMember = false;
                            if (adminUser) {

                              this.props.context.spHttpClient
                                .get(

                                  "/" + this.state.inclusionpath + "/" + this.state.sitename + `/_api/web/lists/GetByTitle('Member List')/Items?$filter= Title eq '${adminUser.Email}'`,
                                  SPHttpClient.configurations.v1
                                )
                                .then((response: SPHttpClientResponse) => {
                                  response.json().then((datada) => {
                                    if (!datada.error) {
                                      let val = datada.value;
                                      if (val.length === 0 && item.displayName === "Member List") {
                                        {
                                          const listDefinition: any = {
                                            Title: adminUser.Email,
                                            FirstName: adminUser.DisplayName.split(
                                              " "
                                            )[0],
                                            LastName: adminUser.DisplayName.split(
                                              " "
                                            )[1],
                                            // "Region": '',
                                            // "Country": '',
                                            Role: "Manager",
                                            Status: "Approved",
                                            Group: "IT Pro",
                                            FocusArea: "All",
                                          };
                                          const spHttpClientOptions: ISPHttpClientOptions = {
                                            body: JSON.stringify(
                                              listDefinition
                                            ),
                                          };
                                          const url: string =
                                            "/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/items";
                                          this.props.context.spHttpClient
                                            .post(
                                              url,
                                              SPHttpClient.configurations
                                                .v1,
                                              spHttpClientOptions
                                            )
                                            .then(
                                              (
                                                newUserResponse: SPHttpClientResponse
                                              ) => {
                                                this.setState({
                                                  isShow: false,
                                                });
                                                if (
                                                  newUserResponse.status ===
                                                  201
                                                ) {
                                                  alert(
                                                    ` ${this.state.loggedinUserName} has been added as a manager to the Champion Management Platform, please refresh the app to complete the setup.`
                                                  );
                                                } else {
                                                  alert(
                                                    "Response status " +
                                                    newUserResponse.status +
                                                    " - " +
                                                    newUserResponse.statusText
                                                  );
                                                }
                                              }
                                            );
                                        }
                                      }
                                    }
                                  });
                                });
                            }

                          });
                      }
                    );
                }, 6000);
                setTimeout(() => {
                  if (item.displayName === "Events List") {
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

                        const url: string =

                          "/" +
                          this.state.inclusionpath +
                          "/" +
                          this.state.sitename +
                          "/_api/web/lists/GetByTitle('Events List')/items";
                        this.props.context.spHttpClient
                          .post(
                            url,
                            SPHttpClient.configurations.v1,
                            spHttpClientOptions
                          )
                          .then(
                            (
                              newUserResponse: SPHttpClientResponse
                            ) => {

                            }
                          );
                      }
                    );
                  }
                }, 5000);
              }
            }
          });
      }).catch((error) => {
        alert(errorMessage + "while creating new list. Below are the details: \n" + JSON.stringify(error));
        console.error("CMP_CLBHome_createNewList_FailedtoCreateList \n", JSON.stringify(error));
      });
  }

  //When app is installed create new site collection and lists if not existing already
  private async createSiteAndLists() {
    console.log(cmpLog + "Checking if site exists already.");
    //Set Variables
    var exSiteId;
    try {
      //Check if CMP site exists        
      await sp.site.exists(rootSiteURL + "/" + this.state.inclusionpath + "/" + this.state.sitename).then((response) => {
        if (response != undefined) {
          //If CMP site does not exist, create the site and lists
          if (!response) {
            flagCheckUserRole = false;
            console.log(cmpLog + "Creating new site collection: '" + this.state.sitename + "'");
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
              .then((siteResponse: SPHttpClientResponse) => {
                //If site is succesfully created
                if (siteResponse.status === 200) {
                  console.log(cmpLog + "Created new site collection: '" + this.state.sitename + "'");
                  siteResponse.json().then((siteData: any) => {
                    if (siteData.SiteId) {
                      exSiteId = siteData.SiteId;
                      this.setState({ siteId: siteData.SiteId }, () => {
                        let isMembersListNotExists = false;
                        console.log(cmpLog + "Creating Lists in new site");
                        //Create 3 lists in the newly created site
                        if (exSiteId) {
                          let lists = [];
                          siteconfig.lists.forEach((item) => {
                            let listColumns = [];
                            item.columns.forEach((element) => {
                              let column;
                              switch (element.type) {
                                case "text":
                                  column = {
                                    name: element.name,
                                    text: {},
                                  };
                                  listColumns.push(column);
                                  break;
                                case "choice":
                                  switch (element.name) {
                                    case "Region":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: [
                                            "Africa",
                                            "Asia",
                                            "Australia / Pacific",
                                            "Europe",
                                            "Middle East",
                                            "North America / Central America / Caribbean",
                                            "South America",
                                          ],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "Country":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: ["INDIA", "USA"],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "Role":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: ["Manager", "Champion"],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "Status":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: ["Approved", "Pending"],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "FocusArea":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: [
                                            "Marketing",
                                            "Teamwork",
                                            "Business Apps",
                                            "Virtual Events",
                                          ],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "Group":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: [
                                            "IT Pro",
                                            "Sales",
                                            "Engineering",
                                          ],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    case "Description":
                                      column = {
                                        name: element.name,
                                        choice: {
                                          allowTextEntry: false,
                                          choices: [
                                            "Event Moderator",
                                            "Office Hours",
                                            "Blogs",
                                            "Training",
                                          ],
                                          displayAs: "dropDownMenu",
                                        },
                                      };
                                      listColumns.push(column);
                                      break;
                                    default:
                                      break;
                                  }
                                  break;
                                case "boolean":
                                  column = {
                                    name: element.name,
                                    boolean: {},
                                  };
                                  listColumns.push(column);
                                  break;
                                case "dateTime":
                                  column = {
                                    name: element.name,
                                    dateTime: {},
                                  };
                                  listColumns.push(column);
                                  break;
                                case "number":
                                  column = {
                                    name: element.name,
                                    number: {},
                                  };
                                  listColumns.push(column);
                                  break;
                                default:
                                  break;
                              }
                            });
                            let list = {
                              displayName: item.listName,
                              columns: listColumns,
                              list: {
                                template: "genericList",
                              },
                            };
                            lists.push(list);
                          });

                          lists.forEach((item) => {
                            let siteId =
                              item.displayName === siteconfig.lists[0].listName
                                ? this.state.siteId // siteconfig.rootSiteId
                                : exSiteId;
                            if (
                              item.displayName === siteconfig.lists[0].listName &&
                              !isMembersListNotExists
                            ) {
                              this.createNewList(siteId, item);

                            } else {
                              this.createNewList(siteId, item);

                            }
                          });
                        }
                      });
                    }
                  });
                  //site is newly created and got 200 response, now create Digital Lib
                  this.createDigitalBadgeLib();
                }

              }).catch((error) => {
                alert(errorMessage + "while creating new site. Below are the details: \n" + JSON.stringify(error));
                console.error("CMP_CLBHome_createSiteAndLists_FailedToCreateSite \n", JSON.stringify(error));
              });

          }//IF END
          //If CMP site already exists create only lists.  
          else {
            //Check if Lists exists already
            this.props.context.spHttpClient
              .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/web/lists/GetByTitle('Member List')/Items", SPHttpClient.configurations.v1)
              .then((responseMemberList: SPHttpClientResponse) => {
                if (responseMemberList.status === 404) {
                  //If lists do not exist create lists. Else no action is required
                  console.log(cmpLog + "Site already existing but lists not found");
                  console.log(cmpLog + "Getting site collection ID for creating lists");
                  flagCheckUserRole = false;
                  //Get Sitecollection ID for creating lists   
                  this.props.context.spHttpClient
                    .get("/" + this.state.inclusionpath + "/" + this.state.sitename + "/_api/site/id", SPHttpClient.configurations.v1)
                    .then((responseuser: SPHttpClientResponse) => {
                      if (responseuser.status === 404) {
                        alert(errorMessage + "while setting up the App. Please try refreshing or loading after some time.");
                        console.error("CMP_CLBHome_createSiteAndLists_FailedToGetSiteID \n");
                      }
                      else {
                        responseuser.json().then((datauser: any) => {
                          exSiteId = datauser.value;
                          if (exSiteId) {
                            console.log(cmpLog + "Creating lists");
                            //Set up List Creation Information
                            let lists = [];
                            siteconfig.lists.forEach((item) => {
                              let listColumns = [];
                              item.columns.forEach((element) => {
                                let column;
                                switch (element.type) {
                                  case "text":
                                    column = {
                                      name: element.name,
                                      text: {},
                                    };
                                    listColumns.push(column);
                                    break;
                                  case "choice":
                                    switch (element.name) {
                                      case "Region":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: [
                                              "Africa",
                                              "Asia",
                                              "Australia / Pacific",
                                              "Europe",
                                              "Middle East",
                                              "North America / Central America / Caribbean",
                                              "South America",
                                            ],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "Country":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: ["INDIA", "USA"],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "Role":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: ["Manager", "Champion"],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "Status":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: ["Approved", "Pending"],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "FocusArea":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: [
                                              "Marketing",
                                              "Teamwork",
                                              "Business Apps",
                                              "Virtual Events",
                                            ],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "Group":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: [
                                              "IT Pro",
                                              "Sales",
                                              "Engineering",
                                            ],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      case "Description":
                                        column = {
                                          name: element.name,
                                          choice: {
                                            allowTextEntry: false,
                                            choices: [
                                              "Event Moderator",
                                              "Office Hours",
                                              "Blogs",
                                              "Training",
                                            ],
                                            displayAs: "dropDownMenu",
                                          },
                                        };
                                        listColumns.push(column);
                                        break;
                                      default:
                                        break;
                                    }
                                    break;
                                  case "boolean":
                                    column = {
                                      name: element.name,
                                      boolean: {},
                                    };
                                    listColumns.push(column);
                                    break;
                                  case "dateTime":
                                    column = {
                                      name: element.name,
                                      dateTime: {},
                                    };
                                    listColumns.push(column);
                                    break;
                                  case "number":
                                    column = {
                                      name: element.name,
                                      number: {},
                                    };
                                    listColumns.push(column);
                                    break;
                                  default:
                                    break;
                                }
                              });
                              let list = {
                                displayName: item.listName,
                                columns: listColumns,
                                list: {
                                  template: "genericList",
                                },
                              };
                              lists.push(list);
                            });

                            //Iterate and create all the lists
                            lists.forEach((item) => {
                              this.createNewList(exSiteId, item);
                            });

                          }

                        });
                      }
                    }).catch((error) => {
                      alert(errorMessage + "while retrieving SiteID. Below are the details: \n" + JSON.stringify(error));
                      console.error("CMP_CLBHome_createSiteAndLists_FailedToGetSiteID \n", JSON.stringify(error));
                    });

                }
              }).catch((error) => {
                alert(errorMessage + "while checking if MemberList exists. Below are the details: \n" + JSON.stringify(error));
                console.error("CMP_CLBHome_createSiteAndLists_FailedToCheckMemberListExists \n", JSON.stringify(error));
              });
            //site exists, check if Digital Badge lib exists, create if not present
            this.createDigitalBadgeLib();

          }//End of else part - CMP site already exists create only lists. 

        }//First IF END

      }).catch((error) => {
        alert(errorMessage + "while checking if site exists. Below are the details: \n" + JSON.stringify(error));
        console.error("CMP_CLBHome_createSiteAndLists_FailedtoCheckIfSiteExists \n", JSON.stringify(error));
      });
    }
    catch (error) {
      console.error("CMP_CLBHome_createSiteAndLists \n", error);
      alert(errorMessage + "while creating site and lists. Below are the details: \n" + error);
    }

  }
  //create digital lib to store the badges
  private async createDigitalBadgeLib() {
    try {
      const listStructure: any = siteconfig.libraries;
      for (let i = 0; i < listStructure.length; i++) {
        const spListTitle: string = listStructure[i]["listName"];
        const spListTemplate = listStructure[i]["listTemplate"];
        //check if digital assests lib exists, create if doesn't exists and upload default badge 
        await spweb.lists.getByTitle(spListTitle).get().then(async () => {

        }).
          catch(async () => {
            //create lib
            await spweb.lists.add(spListTitle, "", spListTemplate, true).then(async () => {
              fetch(require('../assets/images/CMPBadge.png')).then(res => res.blob()).then((blob) => {
                spweb.getFolderByServerRelativeUrl("/" + this.state.inclusionpath + "/" + this.state.sitename + "/" + spListTitle).files.add("digitalbadge.png", blob, true)
                  .then((res) => {
                    res.file.getItem().then(item => {
                      item.update({
                        Title: "Teamwork Champion"
                      });
                    });
                  });
              });
              const deafultXML = await spweb.lists.getByTitle(spListTitle).defaultView.fields.getSchemaXml();
              let titleFieldIndex = deafultXML.indexOf("Title");
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

        console.log(cmpLog + "Checking user role and setting the UI components");
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
            response.json().then((datada) => {
              if (!datada.error) {
                let dataexists: any = datada.value.find(
                  (x) =>
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
          }).catch((error) => {
            alert(errorMessage + "while retrieving user role. Below are the details: \n" + JSON.stringify(error));
            console.error("CMP_CLBHome_checkUserRole_FailedtoGetUserRole \n", JSON.stringify(error));
          });
      }
    }
    catch (error) {
      console.error("CMP_CLBHome_checkUserRole \n", error);
      alert(errorMessage + " while retrieving user role. Below are the details: \n" + error);
    }
  }

  public render(): React.ReactElement<IClbHomeProps> {
    return (
      <div className={styles.clbHome} >
        {this.state.isShow && <div className={styles.load}></div>}
        < div className={styles.container} >
          <div>
            <Header
              showSearch={this.state.cB}
              clickcallback={() =>
                this.setState({
                  cB: false,
                  ChampionsList: false,
                  addMember: false,
                  dB: false,
                  approveMember: false,
                  enableTOT: false
                })
              }
            />
          </div>
          {!this.state.cB &&
            !this.state.ChampionsList &&
            !this.state.addMember && !this.state.approveMember &&
            !this.state.dB && !this.state.enableTOT && (
              <div>
                <div className={styles.imgheader}>
                  <span className={styles.cmpPageHeading}>Welcome {this.state.firstName}!</span>
                </div>
                <div className={styles.grid}>
                  <div className={styles.quickguide}>Get Started</div>
                  <Row className="mt-4">
                    <Col sm={3} className={styles.imageLayout}>
                      <Media
                        className={styles.cursor}
                        onClick={() => this.setState({ cB: !this.state.cB })}
                      >
                        <div className={styles.mb}>
                          <img
                            src={require("../assets/CMPImages/ChampionLeaderBoard.svg")}
                            alt="Champion Leader Board"
                            title="Champion Leader Board"
                            className={styles.dashboardimgs}
                          />
                          <div className={styles.center} title="Champion Leader Board">
                            Champion Leader Board
                          </div>
                        </div>
                      </Media>
                    </Col>
                    {(this.state.cV || this.state.clB) && (
                      <Col sm={3} className={styles.imageLayout}>
                        <Media
                          className={styles.cursor}
                          onClick={() =>
                            this.setState({
                              addMember: !this.state.addMember,
                            })
                          }
                        >
                          <div className={styles.mb}>
                            <img
                              src={require("../assets/CMPImages/AddMembers.svg")}
                              alt="Adding Members Start adding the people you will collaborate with in your..."
                              title="Adding Members Start adding the people you will collaborate with in your..."
                              className={styles.dashboardimgs}
                            />
                            <div className={styles.center} title="Add Members">Add Members</div>
                          </div>
                        </Media>
                      </Col>
                    )}
                    {(this.state.cV || this.state.clB) && (
                      <Col sm={3} className={styles.imageLayout}>
                        <Media
                          className={styles.cursor}
                          onClick={() => this.setState({ dB: !this.state.dB })}
                        >
                          <div className={styles.mb}>
                            <img
                              src={require("../assets/CMPImages/DigitalBadge.svg")}
                              alt="Digital Badge, Get your Champion Badge"
                              title="Digital Badge, Get your Champion Badge"
                              className={styles.dashboardimgs}
                            />
                            <div className={styles.center} title="Digital Badge">Digital Badge</div>
                          </div>
                        </Media>
                      </Col>
                    )}
                    {this.state.isTOTEnabled && (
                      <Col sm={3} className={styles.imageLayout}>
                        <div>
                          <Media
                            className={styles.cursor}
                            onClick={() => this.setState({ enableTOT: !this.state.enableTOT })}
                          >
                            <div className={styles.mb}>
                              <img
                                src={require("../assets/CMPImages/TournamentOfTeams.svg")}
                                alt="Tournament of Teams"
                                title="Tournament of Teams"
                                className={styles.dashboardimgs}
                              />
                              {this.state.isTOTEnabled && (<div className={`${styles.center} ${styles.totLabel}`} title="Tournament of Teams">
                                Tournament of Teams</div>)}
                            </div>
                          </Media>
                        </div>
                      </Col>)}
                  </Row>

                  {this.state.clB && !this.state.cV && (
                    <div className={styles.admintools}>Admin Tools</div>)}

                  {this.state.clB && !this.state.cV && (
                    <Row className="mt-4">
                      <Col sm={3} className={styles.imageLayout}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`/${this.state.inclusionpath}/${this.state.sitename}/Lists/Member%20List/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/CMPImages/ChampionList.svg")}
                                alt="Accessing Champions List"
                                title="Accessing Champions List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center} title="Champions List">Champions List</div>
                          </div>
                        </Media>
                      </Col>
                      <Col sm={3} className={styles.imageLayout}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`/${this.state.inclusionpath}/${this.state.sitename}/Lists/Events%20List/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/CMPImages/EventsList.svg")}
                                alt="Accessing Events List"
                                title="Accessing Events List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center} title="Events List">Events List</div>
                          </div>
                        </Media>
                      </Col>
                      <Col sm={3} className={styles.imageLayout}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`/${this.state.inclusionpath}/${this.state.sitename}/Lists/Event%20Track%20Details/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/CMPImages/EventTrackList.svg")}
                                alt="Accessing Event Track List"
                                title="Accessing Event Track List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center} title="Event Track List">
                              Event Track List
                            </div>
                          </div>
                        </Media>
                      </Col>

                      <Col sm={3} className={styles.imageLayout}>
                        <Media className={styles.cursor}
                          onClick={() =>
                            this.setState({
                              approveMember: !this.state.approveMember,
                            })
                          }
                        >
                          <div className={styles.mb}>
                            <img
                              src={require("../assets/CMPImages/ManageApprovals.svg")}
                              alt="Approve Champion"
                              title="Approve Champion"
                              className={styles.dashboardimgs}
                            />
                            <div className={styles.center} title="Manage Approvals">
                              Manage Approvals
                            </div>
                          </div>
                        </Media>
                      </Col>
                      {!this.state.isTOTEnabled && (
                        <Col sm={3} className={styles.imageLayout}>
                          <Media
                            className={styles.cursor}
                            onClick={() => this.setState({ enableTOT: !this.state.enableTOT })}
                          >
                            <div className={styles.mb}>
                              <img
                                src={require("../assets/CMPImages/EnableTOT.svg")}
                                alt="Enable Tournament of Teams"
                                title="Enable Tournament of Teams"
                                className={styles.dashboardimgs}
                              />
                              {!this.state.isTOTEnabled && (<div className={`${styles.center} ${styles.enableTournamentLabel}`} title="Enable Tournament of Teams">
                                Enable Tournament of Teams</div>)}
                            </div>
                          </Media>
                        </Col>)}
                      <Col sm={3} className={styles.imageLayout}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`/${this.state.inclusionpath}/${this.state.sitename}/Digital%20Badge%20Assets/Forms/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/CMPImages/ManageDigitalBadges.svg")}
                                alt="Manage Digital Badges"
                                title="Manage Digital Badges"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center} title="Manage Digital Badges">
                              Manage Digital Badges
                            </div>
                          </div>
                        </Media>
                      </Col>

                    </Row>
                  )}
                  {this.state.clB && !this.state.cV && (
                    <Row>

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
                onClickCancel={() => this.setState({ clB: true, cB: false })}
              />
            )
          }
          {
            this.state.cB && this.state.cV && (
              <ChampionLeaderBoard
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                onClickCancel={() => this.setState({ clB: false, cB: false })}
              />
            )
          }
          {
            this.state.enableTOT && (
              <TOTLandingPage
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                isTOTEnabled={this.state.isTOTEnabled}
              />
            )
          }
          {
            this.state.addMember && (
              <ClbAddMember
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                onClickCancel={() =>
                  this.setState({ addMember: false, ChampionsList: true, isUserAdded: false })
                }
                onClickSave={(userStatus: string) =>
                  this.setState({ addMember: false, ChampionsList: true, isUserAdded: true, userStatus: userStatus })
                }
                onClickBack={() => { this.setState({ addMember: false }); }}
              />
            )
          }
          {
            this.state.approveMember && (
              <ApproveChampion
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                isEmp={this.state.cV === true || this.state.clB === true}
                onClickAddmember={() =>
                  this.setState({
                    cB: false,
                    ChampionsList: false,
                    addMember: false,
                    approveMember: false
                  })
                }
              />
            )
          }
          {
            this.state.ChampionsList && (
              <ClbChampionsList
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                isEmp={this.state.cV === true || this.state.clB === true}
                userAdded={this.state.isUserAdded}
                userStatus={this.state.userStatus}
                onClickAddmember={() =>
                  this.setState({
                    cB: false,
                    ChampionsList: false,
                    addMember: false,
                    approveMember: false
                  })
                }
              />
            )
          }
          {
            this.state.cB && this.state.eV && (
              <EmployeeView
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                onClickCancel={() => this.setState({ eV: true, cB: false })}
              />
            )
          }
          {
            this.state.dB && (
              <DigitalBadge
                siteUrl={this.props.siteUrl}
                context={this.props.context}
                clientId=""
                description=""
                theme={ThemeStyle.Light}
                fontSize={12}
                clickcallback={() => this.setState({ dB: false })}
                clickcallchampionview={() =>
                  this.setState({ cB: false, eV: false, dB: false })
                }
              />
            )
          }
        </div >
      </div >
    );
  }
}
