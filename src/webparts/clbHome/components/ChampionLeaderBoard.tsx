import * as React from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import "../scss/Championleaderboard.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as _ from "lodash";
import Sidebar from "./Sidebar";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import ChampionsCards from "./ChampionsCards";
import { Label } from "@fluentui/react";
import * as constants from '../constants/strings';

interface ChampionLeaderBoardProps {
  context: WebPartContext;
  onClickCancel: Function;
  siteUrl: string;
  appTitle: string;
  loggedinUserEmail: string;
  currentThemeName?: string;
}

export default function ChampionLeaderBoard(props: ChampionLeaderBoardProps) {
  const [userCalled, setUserCalled] = React.useState(false);
  const [isLoaded, setIsLoaded] = React.useState(false);
  const [isUpdated, setIsUpdated] = React.useState(false);
  const [eventsSubmissionMessage, setEventsSubmissionMessage] = React.useState("");
  const isDarkOrContrastTheme = props.currentThemeName === constants.themeDarkMode || props.currentThemeName === constants.themeContrastMode;

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
            setEventsSubmissionMessage={setEventsSubmissionMessage}
            currentThemeName={props.currentThemeName}
          />
          <div className="content-tab">
            <div className={`ClbPath${isDarkOrContrastTheme ? ' ClbPathDarkContrast' : ""}`}>
              <img src={require("../assets/CMPImages/BackIcon.png")}
                className="backImg"
                alt={LocaleStrings.BackButton}
                aria-hidden="true"
              />
              <span
                className="backLabel"
                onClick={() => { props.onClickCancel(); }}
                role="button"
                tabIndex={0}
                onKeyDown={(evt: any) => { if (evt.key === constants.stringEnter) props.onClickCancel(); }}
                aria-label={props.appTitle}

              >
                <span title={props.appTitle}>
                  {props.appTitle}
                </span>
              </span>
              <span className="ClbBorder"></span>
              <span className="ClbLabel">{LocaleStrings.ChampionLeaderBoardLabel}</span>
            </div>
            {eventsSubmissionMessage !== "" &&
              <Label className={`events-success-message${isDarkOrContrastTheme ? ' events-success-messageDarkContrast' : ""}`}>
                <img src={require('../assets/TOTImages/tickIcon.png')}
                  alt={LocaleStrings.SuccessIcon} className="tickImage" />
                {eventsSubmissionMessage}
              </Label>
            }
            <ChampionsCards
              siteUrl={props.siteUrl}
              context={props.context}
              type={LocaleStrings.PivotHeaderGlobal}
              callBack={_renderListAsync}
              loggedinUserEmail={props.loggedinUserEmail}
              currentThemeName={props.currentThemeName}
            />
          </div>
        </div>
      )}
    </div>
  );
}
