import * as React from "react";
import { Component } from "react";
import "../scss/Championview.scss";
import Row from "react-bootstrap/Row";
import Accordion from "react-bootstrap/Accordion";
import Card from "react-bootstrap/Card";
import Table from "react-bootstrap/Table";
import * as microsoftTeams from "@microsoft/teams-js";
import { ILabelStyles, IStyleSet, Label, Icon, initializeIcons, } from "office-ui-fabric-react";
initializeIcons();

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 20 },
};
interface ChampionsProps {
  users: any;
  type: string;
  fromV: string;
  events?: any;
  filterBy?: string;
  callBack?: Function;
}
interface ChampionState {
  isLoaded: boolean;
}
export default class Champions extends Component<
  ChampionsProps,
  ChampionState
> {
  constructor(props: any) {
    super(props);
    this.state = {
      isLoaded: false,
    };
    this._renderListAsync();
  }
  public _renderListAsync() {
    microsoftTeams.initialize();
    this.setState({ isLoaded: true });
  }

  public componentDidMount() {
    setTimeout(() => {
      this.setState({ isLoaded: true });
    }, 500);
  }
  public openTask = (selectedTask: string) => {
    microsoftTeams.initialize();
    microsoftTeams.executeDeepLink(selectedTask);
  }
  public addDefaultSrc(ev) {
    ev.target.src = require("../assets/images/noprofile.png");
  }

  public render() {
    const starStyles = {
      color: "#f3ca3e"
    };

    return (
      <React.Fragment>
        {this.props.type && (
          <Label className="gtcLabel">            
            <span className="gtc">
              <b>{this.props.type}</b> Top {this.props.filterBy} Champions
            </span>
          </Label>
        )}
        {this.props.users.length === 0 && (
          <div className="m-4 card">
            <b className="card-title p-4 text-center">Records Not Found</b>
          </div>
        )}
        {this.props.users.length > 0 && (
          <>
            <div className="gtc-cards">
              <Row>
                {this.props.users
                  .filter(
                    (user: { Role: string; Status: string }) =>
                      (user.Role === "Champion" || user.Role === "Manager") &&
                      user.Status === "Approved"
                  )
                  .map((member: any, ind = 0) => {
                    return (
                      this.state.isLoaded && (
                        <div className={ind === 0 ? "cards" : "cards"}>
                          {this.props.fromV == "e" && (
                            <img
                              src={require("../assets/CMPImages/badgeStripNoRank.png")}
                              alt="top-badge"
                            />
                          )}
                          {this.props.fromV != "e" && (
                            <img
                              src={require("../assets/CMPImages/badgeStripNew.png")}
                              alt="top-badge"
                            />
                          )}
                          {this.props.fromV != "e" && (
                            <div className="rank">
                              <span>#</span>
                              {ind + 1}
                            </div>
                          )}
                          <img
                            src={
                              "/_layouts/15/userphoto.aspx?username=" +
                              member.Title
                            }
                            className="profile-img"
                            onError={this.addDefaultSrc}
                          />
                          <div className={ind === 0 ? "gtc-name2" : "gtc-name"}>
                            {member.FirstName}
                          </div>
                          {this.props.fromV != "e" && (
                            <>
                              <div className="gtc-star">
                                <Icon
                                  iconName="CannedChat"
                                  className="chat-icon"
                                  onClick={() => this.openTask(`https://teams.microsoft.com/l/chat/0/0?users=${member.Title}`)}
                                />
                                <Icon iconName="FavoriteStarFill"  style={starStyles} id="points" />
                                <span className="totalPoints">{member.totalpoints}</span>
                                <a href={`mailto:${member.Title}`}>
                                  <Icon
                                    iconName="NewMail"
                                    className="mail-icon"
                                  ></Icon>
                                </a>
                              </div>
                            </>
                          )}
                          {this.props.fromV == "e" && (
                            <>
                              <Icon
                                iconName="CannedChat"
                                className="chat-icon"
                                onClick={() =>
                                  this.openTask(
                                    `https://teams.microsoft.com/l/chat/0/0?users=${member.Title}`
                                  )
                                }
                              />
                              <a href={`mailto:${member.Title}`}>
                                <Icon
                                  iconName="NewMail"
                                  className="mail-icon"
                                ></Icon>
                              </a>
                            </>
                          )}
                        </div>
                      )
                    );
                  })}
              </Row>
            </div>
            <div className="paddingTop">
              {this.props.type && (
                <React.Fragment>
                  <span className="gtc">
                    <b>{this.props.type}</b> Top {this.props.filterBy} Champions
                    : <b>My Rank</b>
                  </span>
                  <div className="table-content">
                    {this.state.isLoaded && (
                      <Accordion>
                        {this.props.users
                          .slice(0, 3)
                          .map((rankedMember: any, ind: number) => {
                            return (
                              <Card className="topChampCards">
                                <Accordion.Toggle
                                  as={Card.Header}
                                  eventKey={rankedMember.ID}
                                >
                                  <div className="gttc-row-left">
                                    <span>
                                      <img src={"/_layouts/15/userphoto.aspx?username=" + rankedMember.Title} className="gttc-img" onError={this.addDefaultSrc} />
                                      <div className="gttc-img-name">
                                        {rankedMember.FirstName}
                                      </div>
                                    </span>
                                  </div>                                  
                                    <div className="gttc-row-right">
                                      <div className="gttc-star">
                                        <Icon
                                          iconName="FavoriteStarFill"
                                          id="points2"
                                          style={starStyles}                                         
                                        />
                                        <span className="points">{rankedMember.totalpoints}</span>
                                      </div>
                                      <div className="vline"></div>
                                      <div className="gttc-rank">
                                        Rank <b>{ind + 1}</b>
                                      </div>
                                  </div>
                                </Accordion.Toggle>
                                <Accordion.Collapse eventKey={rankedMember.ID}>
                                  <Card.Body>
                                    <Table>
                                      {Object.keys(rankedMember.eventpoints).length != 0 &&
                                        <tr>
                                          <th>Event Type</th>
                                          <th className="countHeader">Count</th>
                                        </tr>
                                      }
                                      {Object.keys(rankedMember.eventpoints).map((e, i) => {
                                        return (
                                          e !== "0" && (
                                            <tr>
                                              <td className="eventTypeCol">
                                                {
                                                  this.props.events.find((ev) => e !== "0" && ev.ID.toString() === e)
                                                  && this.props.events.find((ev) => e !== "0" && ev.ID.toString() === e).Title
                                                }
                                              </td>
                                              <td className="gttc-tap-data">
                                                {
                                                  rankedMember.eventpoints[e].map((x) => x.Count / x.Count).length
                                                }
                                              </td>
                                            </tr>
                                          )
                                        );
                                      })}
                                    </Table>
                                  </Card.Body>
                                </Accordion.Collapse>
                              </Card>
                            );
                          })}
                      </Accordion>
                    )}
                  </div>
                </React.Fragment>
              )}
            </div>
          </>
        )}
      </React.Fragment>
    );
  }
}
