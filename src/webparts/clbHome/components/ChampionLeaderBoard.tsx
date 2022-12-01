import * as React from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import "../scss/Championleaderboard.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as _ from "lodash";
import Sidebar from "./Sidebar";
import ChampionvView from "./ChampionvView";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import ChampionsCards from "./ChampionsCards";

interface ChampionLeaderBoardProps {
  context: WebPartContext;
  onClickCancel: Function;
  siteUrl: string;
}

export default function ChampionLeaderBoard(props: ChampionLeaderBoardProps) {
  const [userCalled, setUserCalled] = React.useState(false);
  const [isLoaded, setIsLoaded] = React.useState(false);
  const [isUpdated, setIsUpdated] = React.useState(false);

  const _renderListAsync = async () => {
    setIsLoaded(true);
    //Setting state to re-render the child components
    setIsUpdated(!isUpdated);
  };

  React.useEffect(() => {
    if (!userCalled) {
      setUserCalled(true);
      _renderListAsync();
    }
  });

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
          <div className="content-tab">
            <div className="ClbPath">
              <img src={require("../assets/CMPImages/BackIcon.png")}
                className="backImg"
                alt={LocaleStrings.BackButton}
              />
              <span
                className="backLabel"
                onClick={() => { props.onClickCancel(); }}
                title={LocaleStrings.CMPBreadcrumbLabel}
              >
                {LocaleStrings.CMPBreadcrumbLabel}
              </span>
              <span className="ClbBorder"></span>
              <span className="ClbLabel">{LocaleStrings.ChampionLeaderBoardLabel}</span>
            </div>
            <ChampionsCards
              siteUrl={props.siteUrl}
              context={props.context}
              type={LocaleStrings.PivotHeaderGlobal}
              callBack={_renderListAsync}
            />
            <ChampionvView
              siteUrl={props.siteUrl}
              context={props.context}
              callBack={_renderListAsync}
            />
          </div>
        </div>
      )}
    </div>
  );
}
