import * as React from "react";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import styles from "../scss/ClbHome.module.scss";
import { Dropdown, IDropdownStyles } from "office-ui-fabric-react/lib/Dropdown";
import { autobind } from "office-ui-fabric-react/lib/Utilities";
import { sp } from "@pnp/sp";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import Row from "react-bootstrap/Row";
import Col from "react-bootstrap/Col";
import siteconfig from "../config/siteconfig.json";

export interface IClbAddMemberProps {
  context?: any;
  onClickCancel: () => void;
  siteUrl: string;
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
  ID: number;
  LoginName: string;
  Name : string;
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  SuccessMessage: string;
  UserDetails: Array<any>;
  selectedusers: Array<any>;
  siteUrl: string;
  regionDropdown: Array<any>;
  allUser: Array<any>;
  coutries: Array<any>;
  regions: Array<any>;
  users: Array<any>;
  roles: Array<any>;
  status: Array<any>;
  groups: Array<any>;
  focusAreas: Array<any>;
  memberData: any;
  memberrole: string;
  sitename: string;
  inclusionpath: string;
}

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: "auto", margin: "1rem 1rem 0 1rem" },
};

class ClbAddMember extends React.Component<IClbAddMemberProps, IState> {
  constructor(props: IClbAddMemberProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context,
    });
    this.state = {
      list: { value: [] },
      isAddChampion: false,
      SuccessMessage: "",
      UserDetails: [],
      selectedusers: [],
      regionDropdown: [],
      allUser: [],
      coutries: [],
      regions: [],
      users: [],
      roles: [],
      status: [],
      groups: [],
      focusAreas: [],
      memberData: { region: "", group: "", focusArea: "", country: "" },
      siteUrl: this.props.siteUrl,
      memberrole: "",
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
    };
  }

  public componentDidMount() {
    this.props.context.spHttpClient
      .get(
       
        "/"+this.state.inclusionpath+"/"+this.state.sitename+"/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Region')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((regions) => {
          if (!regions.error) {
            this.props.context.spHttpClient
              .get(
               
                "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Country')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line: no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((coutries) => {
                  if (!coutries.error) {
                    this.setState({
                      regions: regions.Choices,
                      coutries: coutries.Choices,
                    });
                  }
                });
              });
          }
        });
      });

    this.props.context.spHttpClient
      .get(
       
        "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Group')",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        response.json().then((groups) => {
          if (!groups.error) {
            this.props.context.spHttpClient
              .get(
               
                "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('FocusArea')",
                SPHttpClient.configurations.v1
              )
              // tslint:disable-next-line: no-shadowed-variable
              .then((response: SPHttpClientResponse) => {
                response.json().then((focusAreas) => {
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
  }

  @autobind
  private _getPeoplePickerItems(items: any[]) {
    let userarr: IUserDetail[] = [];
    items.forEach((user) => {
      userarr.push({ ID: user.id, LoginName: user.loginName, Name: user.text });
    });
    this.setState({ UserDetails: userarr });
  }

  private async _getListData(email: any): Promise<any> {
    return this.props.context.spHttpClient
      .get(
        "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/web/lists/GetByTitle('Member List')/Items?$filter=Title eq '" + email.toLowerCase() +"'",
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
  public async _createorupdateItem() {
    return this.props.context.spHttpClient
      .get(
       
        "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          if (!datauser.error) {
            this.props.context.spHttpClient
              .get(
               
                "/"+this.state.inclusionpath+"/"+this.state.sitename+ "/_api/web/lists/GetByTitle('Member List')/Items",
                SPHttpClient.configurations.v1
              )
              .then((responsen: SPHttpClientResponse) => {
                responsen.json().then((datada) => {
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
                      .get( "/" + this.state.inclusionpath + "/" + this.state.sitename+
                         "/_api/web/siteusers",
                        SPHttpClient.configurations.v1
                      )
                      .then((responseData: SPHttpClientResponse) => {
                        if (responseData.status === 200) {
                          responseData.json().then(async (data) => {
                            // tslint:disable-next-line: no-function-expression
                            var member:any=[];
                            data.value.forEach(element => {
                              if(element.Email.toLowerCase() === email.toLowerCase()) 
                              member.push(element);
                            });

                            const listDefinition: any = {
                              Title: email,
                              FirstName: this.state.UserDetails[0].Name.split(" ")[0],
                              LastName: this.state.UserDetails[0].Name.split(" ")[1],
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
                                this.state.memberData.focusArea || "Teamwork",
                            };
                            const spHttpClientOptions: ISPHttpClientOptions = {
                              body: JSON.stringify(listDefinition),
                            };
                            let flag = await this._getListData(email);
                            if (flag == 0) {
                              const url: string =
                              "/"+this.state.inclusionpath+"/"+this.state.sitename+"/_api/web/lists/GetByTitle('Member List')/items";
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
                                    });
                                    alert("User Added successfully");
                                    this.props.onClickCancel();
                                  } else {
                                    alert(
                                      "Response status " +
                                        response.status +
                                        " - " +
                                        response.statusText
                                    );
                                  }
                                });
                            } else {
                              alert("User Already a Champion");
                            }
                          });
                        } else {
                          alert(
                            "Response status " +
                              responseuser.status +
                              " - " +
                              responseuser.statusText
                          );
                        }
                      });
                  }
                });
              });
          }
        });
      });
  }

  public filterUsers(type: string, value: any) {
    if (value.target.innerText !== "All") {
      this.setState({
        memberData: {
          ...this.state.memberData,
          [type]: value.target.innerText,
        },
      });
    }
  }

  public options = (optionArray: any) => {
    let myoptions = [];
    myoptions.push({ key: "All", text: "All" });
    optionArray.forEach((element: any) => {
      myoptions.push({ key: element, text: element });
    });
    return myoptions;
  }

  public onRenderCaretDown = (): JSX.Element => {
    return <span></span>;
  }

  public render() {
    return (
      <div className={styles.clbHome}>
        <div className="container">
          <PeoplePicker
            context={this.props.context}
            titleText="Members"
            personSelectionLimit={3}
            showtooltip={true}
            required={true}
            onChange={this._getPeoplePickerItems}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={this.state.selectedusers}
            resolveDelay={1000}
          />
          <br></br>
          <Row>
            <Col md={3}>
              <Dropdown
                onChange={(event: any) => this.filterUsers("region", event)}
                placeholder="Select an Region"
                options={this.options(this.state.regions)}
                styles={dropdownStyles}
                onRenderCaretDown={this.onRenderCaretDown}
              />
            </Col>
            <Col md={3}>
              <Dropdown
                onChange={(event: any) => this.filterUsers("country", event)}
                placeholder="Select an Country"
                options={this.options(this.state.coutries)}
                styles={dropdownStyles}
                onRenderCaretDown={this.onRenderCaretDown}
              />
            </Col>
            <Col md={3}>
              <Dropdown
                onChange={(event: any) => this.filterUsers("group", event)}
                placeholder="Select an Group"
                options={this.options(this.state.groups)}
                styles={dropdownStyles}
                onRenderCaretDown={this.onRenderCaretDown}
              />
            </Col>
            <Col md={3}>
              <Dropdown
                onChange={(event: any) => this.filterUsers("focusArea", event)}
                placeholder="Select an Focus Area"
                options={this.options(this.state.focusAreas)}
                styles={dropdownStyles}
                onRenderCaretDown={this.onRenderCaretDown}
              />
            </Col>
          </Row>
          <div style={{ float: "right", marginTop: "1rem" }}>
            <button
              className="btn btn-success mr-2"
              onClick={() => this._createorupdateItem()}
            >
              Save
            </button>
            <button
              className="btn btn-secondary"
              onClick={() => this.props.onClickCancel()}
            >
              Cancel
            </button>
          </div>
          <br></br>
          <br></br>
          <label>{this.state.SuccessMessage}</label>
        </div>
      </div>
    );
  }
}

export default ClbAddMember;
