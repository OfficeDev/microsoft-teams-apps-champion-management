import * as React from "react";
import styles from "../scss/ClbHome.module.scss";
import "bootstrap/dist/css/bootstrap.min.css";
import { IClbHomeProps } from "./IClbHomeProps";
import Header from "./Header";
import ChampionLeaderBoard from "./ChampionLeaderBoard";
import ClbAddMember from "./ClbAddMember";
import ClbChampionsList from "./ClbChampionsList";
import EmployeeView from "./EmployeeView";
import Media from "react-bootstrap/Media";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { initializeIcons } from "@uifabric/icons";
import DigitalBadge from "./DigitalBadge";
import { ThemeStyle } from "msteams-ui-styles-core";
import siteconfig from "../config/siteconfig.json";
import { MSGraphClient } from "@microsoft/sp-http";
initializeIcons();

export interface IClbHomeState {
  cB: boolean;
  clB: boolean;
  addMember: boolean;
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
}

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
    };

    this._getListData = this._getListData.bind(this);
    this.rootSiteId = this.rootSiteId.bind(this);
  }

  public componentDidMount() {
    this.setState({
      isShow: true,
    });
    this.rootSiteId();

    this.props.context.spHttpClient
      .get(
        this.state.siteUrl +
          "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          this.setState({ loggedinUserName: datauser.DisplayName });
        });
      });
  }

  //create lists when you upload package into new tenant.

  private _createList() {
    let listname = siteconfig.lists[1].listName;
    const getListUrl: string =
      this.state.siteUrl +
      "/" +
      this.state.inclusionpath +
      "/" +
      this.state.sitename +
      `/_api/web/lists/GetByTitle('${listname}')/Items`;

    let memberListName = siteconfig.lists[0].listName;

    let getMemberListUrl =
      this.state.siteUrl +
      "/" +
      `/_api/web/lists/GetByTitle('${memberListName}')/Items`;

    let isMembersListNotExists = false;
    this.props.context.spHttpClient
      .get(getMemberListUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 404) {
          isMembersListNotExists = true;
        }
      });

    this.props.context.spHttpClient
      .get(getListUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
        } else {
          if (response.status === 404) {
            this.props.context.spHttpClient
              .get(
                this.state.siteUrl +
                  "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                SPHttpClient.configurations.v1
              )
              .then((responseuser: SPHttpClientResponse) => {
                responseuser.json().then((datauser: any) => {
                  const createsiteUrl: string =
                    this.state.siteUrl + "/_api/SPSiteManager/create";
                  const siteDefinition: any = {
                    request: {
                      Title: this.state.sitename,
                      Url:
                        this.state.siteUrl +
                        "/" +
                        this.state.inclusionpath +
                        "/" +
                        this.state.sitename,
                      Lcid: 1033,
                      ShareByEmailEnabled: true,
                      Description: "Description",
                      WebTemplate: "STS#3",
                      SiteDesignId: "6142d2a0-63a5-4ba0-aede-d9fefca2c767",
                      Owner: datauser.Email,
                    },
                  };
                  const spHttpsiteClientOptions: ISPHttpClientOptions = {
                    body: JSON.stringify(siteDefinition),
                  };
                  this.props.context.spHttpClient
                    .post(
                      createsiteUrl,
                      SPHttpClient.configurations.v1,
                      spHttpsiteClientOptions
                    )
                    .then((siteresponse: SPHttpClientResponse) => {
                      if (siteresponse.status === 200) {
                        siteresponse.json().then((sitedata: any) => {
                          if (sitedata.SiteId) {
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
                                item.displayName === memberListName
                                  ? this.state.siteId // siteconfig.rootSiteId
                                  : sitedata.SiteId;
                              if (
                                item.displayName === memberListName &&
                                isMembersListNotExists
                              ) {
                                this.createNewList(siteId, item);
                                setTimeout(() => {
                                  this.props.context.spHttpClient
                                    .get(
                                      this.state.siteUrl +
                                        "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
                                      SPHttpClient.configurations.v1
                                    )
                                    .then(
                                      (
                                        userProperties: SPHttpClientResponse
                                      ) => {
                                        userProperties
                                          .json()
                                          .then(async (adminUser: any) => {
                                            if (adminUser) {
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
                                                this.state.siteUrl +
                                                "/_api/web/lists/GetByTitle('Member List')/items";
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
                                          });
                                      }
                                    );
                                }, 6000);
                              } else {
                               this.createNewList(siteId, item);
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
                                          this.state.siteUrl +
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
                                              if (
                                                newUserResponse.status === 201
                                              ) {
                                                this.props.context.spHttpClient
                                                .post(                                                   
                                                  this.state.siteUrl +
                                                  "/" +
                                                  this.state.inclusionpath +
                                                  "/" +
                                                  this.state.sitename +
                                                  "/_api/web/lists/GetByTitle('Events List')/Items/Fields/GetByTitle('Title')",
                                                  SPHttpClient.configurations.v1,
                                                  {
                                                    headers: {
                                                    // IF-MATCH header: Provides a way to verify that the object being changed has not been changed since it was last retrieved.
                     // "IF-MATCH":"*", will overwrite any modification in the object, since it was last retrieved.
                     "IF-MATCH": "*",
                     "X-HTTP-Method": "PATCH",
                     // Accept header: Specifies the format for response data from the server.
                     "Accept": "application/json;odata=verbose",
                     //Content-Type header: Specifies the format of the data that the client is sending to the server
                     "Content-Type": "application/json;odata=verbose",
                     
                                                    },
                                                    body:  JSON.stringify({
                                                      '__metadata': {
                                                        // Type that you are modifying.
                                                        'type': 'SP.FieldText'
                                                    },
                                     
                                                    'EnforceUniqueValues': true
                                                  }),
                                                  }
                                                )  .then(
                                                  (
                                                    _newUserResponse1: SPHttpClientResponse
                                                  ) => { 
                                                    this.props.context.spHttpClient
                                                    .post(                                                   
                                                      this.state.siteUrl +
                                                      "/" +
                                                      this.state.inclusionpath +
                                                      "/" +
                                                      this.state.sitename +
                                                      "/_api/web/lists/GetByTitle('Events List')/Items/Fields/GetByTitle('Title')",
                                                      SPHttpClient.configurations.v1,
                                                      {
                                                        headers: {
                                                        // IF-MATCH header: Provides a way to verify that the object being changed has not been changed since it was last retrieved.
                         // "IF-MATCH":"*", will overwrite any modification in the object, since it was last retrieved.
                         "IF-MATCH": "*",
                         "X-HTTP-Method": "PATCH",
                         // Accept header: Specifies the format for response data from the server.
                         "Accept": "application/json;odata=verbose",
                         //Content-Type header: Specifies the format of the data that the client is sending to the server
                         "Content-Type": "application/json;odata=verbose",
                         
                                                        },
                                                        body:  JSON.stringify({
                                                          '__metadata': {
                                                            // Type that you are modifying.
                                                            'type': 'SP.FieldText'
                                                        },
                                         
                                                        'EnforceUniqueValues': true
                                                      }),
                                                      }
                                                    )  .then(
                                                      (
                                                        _newUserResponse2: SPHttpClientResponse
                                                      ) => { 
                                                      });
                                                  })
                                                  .catch(
                                                    (
                                                      _newUserResponse2: any
                                                    ) => { 
                                                    });
                                              } else {
                                              }
                                            }
                                          );
                                      }
                                    );
                                  }
                                }, 5000);
                              }
                            });
                          } else {
                            alert("Check your site does not exists already.");
                          }
                        });
                      } else {
                        alert("Check your site does not exists already.");
                      }
                    });
                });
              });
          } else {
            alert(
              "Something went wrong. " +
                response.status +
                " " +
                response.statusText
            );
          }
        }
      });
  }

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
              }
            }
          });
      });
  }

  private rootSiteId() {
    let graphSiteRoot = "sites/root";
    this.props.context.msGraphClientFactory
      .getClient()
      .then((garphClient: MSGraphClient) => {
        garphClient
          .api(graphSiteRoot)
          .version("v1.0")
          .header("Content-Type", "application/json")
          .responseType("json")
          .get()
          .then((data: any) => {
            this.setState({ siteId: data.id.split(",")[1] }, () => {
              this._createList();
            });
          });
      });
  }

  private async _getListData(email: any): Promise<any> {
    return this.props.context.spHttpClient
      .get(
        "/_api/web/lists/GetByTitle('Member List')/Items",
        SPHttpClient.configurations.v1
      )
      .then(async (response: SPHttpClientResponse) => {
        if (response.status === 200) {
          this.setState({
            isShow: false,
          });
          let flag = 0;
          await response.json().then((responseJSON: any) => {
            let i = 0;
            while (i < responseJSON.value.length) {
              if (
                responseJSON.value[i].Title.toLowerCase() == email.toLowerCase()
              ) {
                flag = 1;
                return flag;
              }
              i++;
            }
            return flag;
          });
          return flag;
        }
      });
  }

  public componentWillMount() {
    this.props.context.spHttpClient
      .get(
        "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then(async (datauser) => {
          if (!datauser.error) {
            let flag = await this._getListData(datauser.Email.toLowerCase());
            if (flag === 0) {
              this.setState({ eV: true });
            }
            this.props.context.spHttpClient
              .get(
                this.state.siteUrl +
                  "/_api/web/lists/GetByTitle('Member List')/Items",
                SPHttpClient.configurations.v1
              )
              .then((response: SPHttpClientResponse) => {
                response.json().then((datada) => {
                  if (!datada.error) {
                    let dataexists: any = datada.value.find(
                      (x) =>
                        x.Title.toLowerCase() === datauser.Email.toLowerCase()
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
                  }
                });
              });
          }
        });
      });
  }

  public render(): React.ReactElement<IClbHomeProps> {
    return (
      <div className={styles.clbHome}>
        {this.state.isShow && <div className={styles.load}></div>}
        <div className={styles.container}>
          <div>
            <Header
              showSearch={this.state.cB}
              clickcallback={() =>
                this.setState({
                  cB: false,
                  ChampionsList: false,
                  addMember: false,
                  dB: false,
                })
              }
            />
          </div>
          {!this.state.cB &&
            !this.state.ChampionsList &&
            !this.state.addMember &&
            !this.state.dB && (
              <div>
                <div className={styles.imgheader}>
                  <h2>Welcome to the Microsoft 365</h2>
                  <h3>Champion Management Platform</h3>
                </div>
                <div className={styles.box}></div>
                <div className={styles.grid}>
                  <div className={styles.quickguide}>Quick Start Guide</div>
                  <Row>
                    <Col sm={4}>
                      <Media
                        className={styles.cursor}
                        onClick={() => this.setState({ cB: !this.state.cB })}
                      >
                        <div className={styles.mb}>
                          <img
                            src={require("../assets/images/leaderboard.jpg")}
                            alt="Champion Leader Board"
                            title="Champion Leader Board"
                            className={styles.dashboardimgs}
                          />
                          <div className={styles.center}>
                            Champion Leader Board
                          </div>
                        </div>
                      </Media>
                    </Col>
                    {(this.state.cV || this.state.clB) && (
                      <Col sm={4}>
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
                              src={require("../assets/images/addMember.png")}
                              alt="Adding Members Start adding the people you will collaborate with in your..."
                              title="Adding Members Start adding the people you will collaborate with in your..."
                              className={styles.dashboardimgs}
                            />
                            <div className={styles.center}>Add Members</div>
                          </div>
                        </Media>
                      </Col>
                    )}
                    {(this.state.cV || this.state.clB) && (
                      <Col sm={4}>
                        <Media
                          className={styles.cursor}
                          onClick={() => this.setState({ dB: !this.state.dB })}
                        >
                          <div className={styles.mb}>
                            <img
                              src={require("../assets/images/digitalbadge.png")}
                              alt="Digital Badge, Get your Champion Badge"
                              title="Digital Badge, Get your Champion Badge"
                              className={styles.dashboardimgs}
                            />
                            <div className={styles.center}>Digital Badge</div>
                          </div>
                        </Media>
                      </Col>
                    )}
                  </Row>
                  {this.state.clB && !this.state.cV && (
                    <Row className="mt-4">
                      <Col sm={4}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`${this.state.siteUrl}/Lists/Member%20List/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/images/list.jpg")}
                                alt="Accessing Champions List"
                                title="Accessing Champions List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center}>Champions List</div>
                          </div>
                        </Media>
                      </Col>
                      <Col sm={4}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`${this.state.siteUrl}/${this.state.inclusionpath}/${this.state.sitename}/Lists/Events%20List/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/images/list.jpg")}
                                alt="Accessing Events List"
                                title="Accessing Events List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center}>Events List</div>
                          </div>
                        </Media>
                      </Col>
                      <Col sm={4}>
                        <Media className={styles.cursor}>
                          <div className={styles.mb}>
                            <a
                              href={`${this.state.siteUrl}/${this.state.inclusionpath}/${this.state.sitename}/Lists/Event%20Track%20Details/AllItems.aspx`}
                              target="_blank"
                            >
                              <img
                                src={require("../assets/images/list.jpg")}
                                alt="Accessing Event Track List"
                                title="Accessing Event Track List"
                                className={styles.dashboardimgs}
                              />
                            </a>
                            <div className={styles.center}>
                              Event Track List
                            </div>
                          </div>
                        </Media>
                      </Col>
                    </Row>
                  )}
                </div>
              </div>
            )}
          {this.state.cB && this.state.clB && (
            <ChampionLeaderBoard
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              onClickCancel={() => this.setState({ clB: true, cB: false })}
            />
          )}
          {this.state.cB && this.state.cV && (
            <ChampionLeaderBoard
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              onClickCancel={() => this.setState({ clB: false, cB: false })}
            />
          )}
          {this.state.addMember && (
            <ClbAddMember
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              onClickCancel={() =>
                this.setState({ addMember: false, ChampionsList: true })
              }
            />
          )}
          {this.state.ChampionsList && (
            <ClbChampionsList
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              isEmp={this.state.cV === true || this.state.clB === true}
              onClickAddmember={() =>
                this.setState({
                  cB: false,
                  ChampionsList: false,
                  addMember: false,
                })
              }
            />
          )}
          {this.state.cB && this.state.eV && (
            <EmployeeView
              siteUrl={this.props.siteUrl}
              context={this.props.context}
              onClickCancel={() => this.setState({ eV: true, cB: false })}
            />
          )}
          {this.state.dB && (
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
          )}
        </div>
      </div>
    );
  }
}
