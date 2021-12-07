import * as React from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import "../scss/Championleaderboard.scss";
import {
  IPivotStyles,
  Pivot,
  PivotItem,
} from "office-ui-fabric-react/lib/Pivot";
import { Dropdown, IDropdownStyles } from "office-ui-fabric-react/lib/Dropdown";
import { IStyleSet } from "office-ui-fabric-react/lib/Styling";
import Champions from "../components/Champions";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as _ from "lodash";
import Sidebar from "./Sidebar";
import ChampionvView from "./ChampionvView";
import siteconfig from "../config/siteconfig.json";

type NewType = IPivotStyles;

const pivotStyles: Partial<IStyleSet<NewType>> = {
  link: {
    width: "calc(100% - 66%);",
  },
  linkIsSelected: {
    width: "calc(100% - 70%);",
  },
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: "auto", margin: "1rem 1rem 0 1rem" },
};

interface ChampionLeaderBoardProps {
  context: WebPartContext;
  onClickCancel: Function;
  siteUrl: string;
}

export default function ChampionLeaderBoard(props: ChampionLeaderBoardProps) {
  let data: any;
  const [siteUrl, setSiteUrl] = React.useState(props.siteUrl);
  const [userCalled, setUserCalled] = React.useState(false);
  const [users, setUsers] = React.useState(data);
  const [allUsers, setAllUsers] = React.useState(data);
  const [isLoaded, setIsLoaded] = React.useState(false);
  const [regionDropdown, setRegionDropDown] = React.useState(data);
  const [countryDropdown, setCountryDropDown] = React.useState(data);
  const [eventDropdown, setEventDropDown] = React.useState(data);
  const [filterByFocusArea, setfilterByFocusArea] = React.useState("");
  const [filterBySpeciality, setfilterBySpeciality] = React.useState("");
  const [siteName, setSiteName] = React.useState(siteconfig.sitename);
  const [inclusionpath, setInclusionpath] = React.useState(
    siteconfig.inclusionPath
  );

  const _renderListAsync = async () => {
    props.context.spHttpClient
      .get( "/" + inclusionpath + "/" + siteName + "/_api/web/lists/GetByTitle('Events List')/Items", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response
          .json()
          .then((eventdata) => {
            if (!eventdata.error) {
              setEventDropDown(_.orderBy(eventdata.value.filter(ed => ed.IsActive), ['Id'], ['asc']));
               props.context.spHttpClient.get( 
                "/"+inclusionpath+"/"+siteName+ "/_api/web/lists/GetByTitle('Member List')/Items?$top=1000&$filter= Status eq 'Approved'", SPHttpClient.configurations.v1)
                // tslint:disable-next-line: no-shadowed-variable
                .then((response: SPHttpClientResponse) => {
                  response.json().then((datada) => {
                    if (!datada.error) {
                      let results = processUsers(datada.value);
                      return results;
                    }
                  });
                });
            }
          })
          .catch((e) => { });
      });
  };

  React.useEffect(() => {
    if (!userCalled) {
      setUserCalled(true);
      _renderListAsync();
    }
  });

  const getUserPoints = (id: any) => {
    return props.context.spHttpClient.get(
     
      "/" +
      inclusionpath +
      "/" +
      siteName +
      "/_api/web/lists/GetByTitle('Event Track Details')/Items?$filter=MemberId eq " +
      id,
      SPHttpClient.configurations.v1
    );
  };

  async function processUsers(usersd: any) {
    let result: any[];
    let promises = [];
    for (let i = 0; i < usersd.length; i++) {
      promises.push(getUserPoints(usersd[i].ID));
    }
    result = await Promise.all(promises);
    let c = 0;
    for (let i = 0; i < usersd.length; i++) {
      result[i].json().then((datau) => {
        let eventpoints = _.groupBy(_.orderBy(datau.value, ['Id'], ['asc']), "EventId");
        let pointsTotal = 0;
        if (datau != "undefined") {
          for (let j = 0; j < datau.value.length; j++) {
            pointsTotal += datau.value[j].Count;
          }
        }
        usersd[i]["eventpoints"] = eventpoints;
        usersd[i]["totalpoints"] = pointsTotal;
        c = c + 1;

        if (c === usersd.length) {
          props.context.spHttpClient
            .get(
            
              "/"+inclusionpath+"/"+siteName+ 
              "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('Region')",
              SPHttpClient.configurations.v1
            )
            .then((response: SPHttpClientResponse) => {
              response.json().then((focusAreas) => {
                if (!focusAreas.error) {
                  props.context.spHttpClient
                    .get(
                     
                      "/"+
                      inclusionpath+"/"+siteName+ 
                      "/_api/web/lists/GetByTitle('Member List')/fields/GetByInternalNameOrTitle('FocusArea')",
                      SPHttpClient.configurations.v1
                    )
                    // tslint:disable-next-line: no-shadowed-variable
                    .then((response: SPHttpClientResponse) => {
                      response.json().then((groupList) => {
                        if (!groupList.error) {
                          setRegionDropDown(focusAreas.Choices);
                          setCountryDropDown(groupList.Choices);
                          setUsers(
                            usersd.sort(
                              (
                                a: { totalpoints: number },
                                b: { totalpoints: number }
                              ) => {
                                return b.totalpoints - a.totalpoints;
                              }
                            )
                          );
                          setAllUsers(
                            usersd.sort(
                              (
                                a: { totalpoints: number },
                                b: { totalpoints: number }
                              ) => {
                                return b.totalpoints - a.totalpoints;
                              }
                            )
                          );
                          setIsLoaded(true);
                        }
                      });
                    });
                }
              });
            });
        }
      });
    }
  }


  const filterUsers = (type: string, value: any) => {
    let myUsers = [...allUsers];
    if (value.target.innerText !== "All") {
      myUsers = myUsers.filter((u) => u[type] === value.target.innerText);
      if (type == "Region") {
        setfilterByFocusArea(value.target.innerText);
      } else if (type == "FocusArea") {
        setfilterBySpeciality(value.target.innerText);
      }
    } else {
      setfilterByFocusArea("All");
      setfilterBySpeciality("All");
    }
    setUsers(
      myUsers.sort((a: { totalpoints: number }, b: { totalpoints: number }) => {
        return b.totalpoints - a.totalpoints;
      })
    );
  };

  const options = (optionArray: any) => {
    let myOptions = [];
    myOptions.push({ key: "All", text: "All" });
    optionArray.forEach((element: any) => {
      myOptions.push({ key: element, text: element });
    });
    return myOptions;
  };

  const onRenderCaretDown = (): JSX.Element => {
    return <span></span>;
  };

  return (
    <div>
      {isLoaded && <div className="loader"></div>}
      {isLoaded && (
        <div className="Championleaderboard d-flex">
          <Sidebar
            siteUrl={props.siteUrl}
            context={props.context}
            becomec={false}
            onClickCancel={() => props.onClickCancel()}
            callBack={_renderListAsync}
          />
          <div className="p-0 content-tab">
            <Pivot
              aria-label=""
              styles={pivotStyles}
              onLinkClick={() => {
                setUsers(allUsers);
              }}
              className="pivotControl"
            >
              <PivotItem
                headerText="Global"
                headerButtonProps={{
                  "data-order": 1,
                  "data-title": "Global Data",
                }}
              >
                <Champions
                  users={users}
                  type="Global"
                  events={eventDropdown}
                  fromV={""}
                  filterBy=""
                  callBack={_renderListAsync}
                />
              </PivotItem>
              <PivotItem headerText="Near Me">
                <Dropdown
                  onChange={(event: any) => filterUsers("Region", event)}
                  placeholder="Select Near Me"
                  options={options(regionDropdown)}
                  styles={dropdownStyles}
                  onRenderCaretDown={onRenderCaretDown}
                />
                <Champions
                  users={users}
                  type="Near Me"
                  events={eventDropdown}
                  fromV={""}
                  filterBy={filterByFocusArea}                  
                  callBack={_renderListAsync}                  
                />
              </PivotItem>
              <PivotItem headerText="By Specialty">
                <Dropdown
                  onChange={(event: any) => filterUsers("FocusArea", event)}
                  placeholder="Select By Specialty"
                  options={options(countryDropdown)}
                  styles={dropdownStyles}
                  onRenderCaretDown={onRenderCaretDown}
                />
                <Champions
                  users={users}
                  type="By Specialty"
                  events={eventDropdown}
                  fromV={""}
                  filterBy={filterBySpeciality}                  
                  callBack={_renderListAsync}
                />
              </PivotItem>
            </Pivot>
            <ChampionvView
              siteUrl={props.siteUrl}
              context={props.context}
              callBack={_renderListAsync}
              onClickCancel={() => this.setState({ cB: false, cV: true })}
            />
          </div>
        </div>
      )}
    </div>
  );
}
