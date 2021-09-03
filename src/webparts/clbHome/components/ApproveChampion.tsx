import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, } from "@microsoft/sp-http";
import siteconfig from "../config/siteconfig.json";
export interface IClbChampionsListProps {
  context?: WebPartContext;
  onClickAddmember: Function;
  isEmp: boolean;
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
  FocusArea: String;
  Group: String;
  Role: String;
  Region: string;
  Points: number;
  ID: number;
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  SuccessMessage: string;
  UserDetails: Array<any>;
  selectedusers: Array<any>;
  siteUrl: string;
  memberrole: string;
}
class ApproveChampion extends React.Component<IClbChampionsListProps, IState> {
  constructor(props: IClbChampionsListProps) {
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
      siteUrl: this.props.siteUrl,
      memberrole: "",
    };
    this._getListData();
  }
  
  //Get the list of Members from member List
  private _getListData(): Promise<ISPLists> {
    return this.props.context.spHttpClient
      .get(
         "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          response.json().then((responseJSON: any) => {
            this._renderList(responseJSON.value);
          });
          return response.json();
        }
      });
  }

  private _renderList(items: ISPList[]): void {
    this.setState({ list: { value: items } });
  }

  private updateItem = (e, ID: number) => {
    let ButtonText = e.target.outerText;
    let status = "";
    let Id = ID;
    if (ButtonText === "Approve") {
      status = "Approved";
    }
    else {
      status = "Rejected";
    }
    const listDefinition: any = {
      Status: status,
    };
    const spHttpClientOptions: ISPHttpClientOptions = {
      body: JSON.stringify(listDefinition),
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
    };

    const url: string =
       "/" + siteconfig.inclusionPath + "/" + siteconfig.sitename + `/_api/web/lists/GetByTitle('Member List')/items(${Id})`;
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
          alert("Champion" + status);
        } else {


          alert(
            "Response status " +
            response.status +
            " - " +
            `Champion ${status}.`
          );
          this._getListData();
        }
      });
  }

  public render() {
    return (
      <div>
        <h4 className="mt-2 mb-2">Champion List</h4>
        <table className="table table-bodered table-striped">
          <thead>
            <th>People Name</th>
            <th>Region</th>
            <th>Country</th>
            <th>FocusArea</th>
            <th>Group</th>
            {!this.props.isEmp && <th>Status</th>}
            <th></th>
            <th></th>
          </thead>
          <tbody>
            {this.state.list &&
              this.state.list.value &&
              this.state.list.value.length > 0 &&
              this.state.list.value.map((item: ISPList) => {
                if (item.Status != "Approved" && item.Status != "Rejected") {//showing only approved list
                  return (
                    <tr>
                      <td>
                        {item.FirstName}
                        <span className="mr-1"></span>
                        {item.LastName}
                      </td>
                      <td>{item.Region}</td>
                      <td>{item.Country}</td>
                      <td>{item.FocusArea}</td>
                      <td>{item.Group}</td>
                      {!this.props.isEmp && <td>{item.Status}</td>}
                      <td>
                        <button
                          className="addchampion btn btn-primary"
                          onClick={e => this.updateItem(e, item.ID)}
                        >
                          Approve
                        </button></td>
                      <td>
                        <button
                          className="addchampion btn btn-primary"
                          onClick={e => this.updateItem(e, item.ID)}
                        >
                          Reject
                        </button></td>
                    </tr>
                  );
                }
              })}
            {this.state.list &&
              this.state.list.value &&
              this.state.list.value.length > 0 &&
              this.state.list.value.filter(i => i.Status == "Pending").length == 0 &&
              (
                <tr>
                  <td colSpan={7}>
                    <h5>No champions requests available.</h5>
                  </td>
                </tr>
              )
            }
          </tbody>
        </table>
        <button
          className="addchampion btn btn-primary"
          onClick={() => this.props.onClickAddmember()}
        >
          Back
        </button>
      </div>
    );
  }
}

export default ApproveChampion;
