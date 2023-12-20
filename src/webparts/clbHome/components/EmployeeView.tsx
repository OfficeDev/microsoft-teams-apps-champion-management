import ChampionsCards from './ChampionsCards';
import React, { Component } from 'react';
import Sidebar from '../components/Sidebar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '../scss/Employeeview.scss';
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as strings from "../constants/strings";

interface EmployeeViewState {
}

interface EmployeeViewProps {
  context: WebPartContext;
  onClickCancel: () => void;
  siteUrl: string;
  appTitle: string;
  currentThemeName?: string;
}

export default class EmployeeView extends Component<
  EmployeeViewProps,
  EmployeeViewState
> {
  constructor(props: any) {
    super(props);
  }

  public render() {
    const isDarkOrContrastTheme = this.props.currentThemeName === strings.themeDarkMode || this.props.currentThemeName === strings.themeContrastMode;
    return (
      <div className={`Employeeview d-flex${isDarkOrContrastTheme ? " EmployeeviewDarkContrast" : ""}`}>
        <Sidebar
          siteUrl={this.props.siteUrl}
          context={this.props.context}
          becomec={true}
          onClickCancel={() => this.props.onClickCancel()}
        />
        <div className="main">
          <div className="ClbPath">
            <img src={require("../assets/CMPImages/BackIcon.png")}
              className="backImg"
              alt={LocaleStrings.BackButton}
              aria-hidden="true"
            />
            <span
              className="backLabel"
              onClick={() => { this.props.onClickCancel(); }}
              role="button"
              tabIndex={0}
              onKeyDown={(evt: any) => { if (evt.key === strings.stringEnter) this.props.onClickCancel(); }}
              aria-label={this.props.appTitle}
            >
              <span title={this.props.appTitle}>
                {this.props.appTitle}
              </span>
            </span>
            <span className="ClbBorder"></span>
            <span className="ClbLabel">{LocaleStrings.ChampionLeaderBoardLabel}</span>
          </div>
          <ChampionsCards
            siteUrl={this.props.siteUrl}
            context={this.props.context}
            type={""}
            currentThemeName={this.props.currentThemeName}
          />
        </div>
      </div>
    );
  }
}
