import ChampionsCards from './ChampionsCards';
import React, { Component } from 'react';
import Sidebar from '../components/Sidebar';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '../scss/Employeeview.scss';
import * as LocaleStrings from 'ClbHomeWebPartStrings';


interface EmployeeViewState {
}

interface EmployeeViewProps {
  context: WebPartContext;
  onClickCancel: () => void;
  siteUrl: string;
}

export default class EmployeeView extends Component<
  EmployeeViewProps,
  EmployeeViewState
> {
  constructor(props: any) {
    super(props);
  }

  public render() {
    return (
      <div className="Employeeview d-flex">
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
            />
            <span
              className="backLabel"
              onClick={() => { this.props.onClickCancel(); }}
              title={LocaleStrings.CMPBreadcrumbLabel}
            >
              {LocaleStrings.CMPBreadcrumbLabel}
            </span>
            <span className="ClbBorder"></span>
            <span className="ClbLabel">{LocaleStrings.ChampionLeaderBoardLabel}</span>
          </div>
          <ChampionsCards
            siteUrl={this.props.siteUrl}
            context={this.props.context}
            type={""}
          />
        </div>
      </div>
    );
  }
}
