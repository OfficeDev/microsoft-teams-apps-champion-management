import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
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
}
interface IState {
  list: ISPLists;
  isAddChampion: boolean;
  SuccessMessage: string;
  UserDetails: Array<any>;
  selectedusers: Array<any>;
  siteUrl: string;
  inclusionpath: string;
  sitename : string;
            
}
class ClbChampionsList extends React.Component<IClbChampionsListProps, IState> {
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
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,  
    };
    this._getListData();
  }

  //Get Details of all members from Member List 
  private _getListData(): Promise<ISPLists> {
    return this.props.context.spHttpClient
      .get(  "/"+this.state.inclusionpath+"/"+this.state.sitename+ 
            
        "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000",
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
          </thead>
          <tbody>
            {this.state.list &&
              this.state.list.value &&
              this.state.list.value.length > 0 &&
              this.state.list.value.map((item: ISPList) => {
                if (item.Status === "Approved") {//showing only approved list
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
                    </tr>
                  );
                }
              })}
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

export default ClbChampionsList;
