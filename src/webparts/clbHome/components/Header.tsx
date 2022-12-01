import { DirectionalHint, TooltipHost } from '@fluentui/react';
import "bootstrap/dist/css/bootstrap.min.css";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { Callout, Link, Text } from 'office-ui-fabric-react';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import React, { Component } from "react";
import Navbar from "react-bootstrap/Navbar";
import * as Strings from '../constants/strings';
import styles from "../scss/CMPHome.module.scss";

// Import package version
const packageSolution: any = require("../../../../config/package-solution.json");
const calloutProps = { gapSpace: 0 };
interface IHeaderProps {
  showSearch: boolean;
  clickcallback: () => void; //will redirects to home
  logoImageURL: string;
}
interface HeaderState {
  isCalloutVisible: boolean;

}
export default class Header extends Component<IHeaderProps, HeaderState> {
  constructor(_props: any) {
    super(_props);
    this.state = {
      isCalloutVisible: false,
    };
    this.homeRedirect = this.homeRedirect.bind(this);
  }
  public homeRedirect() {
    this.props.clickcallback();
  }
  public toggleIsCalloutVisible = () => {
    this.setState({ isCalloutVisible: !this.state.isCalloutVisible });
  }
  public render() {
    const buttonId = 'callout-button';
    const labelId = 'callout-label';
    const descriptionId = 'callout-description';
    return (
      <Navbar className={styles.navbg}>
        <Navbar.Brand href="#home" className={styles.white}>
          <img
            src={this.props.logoImageURL}
            className={`d-inline-block ${styles.clbHeaderLogo}`}
            alt="applogo"
            onClick={this.homeRedirect}
            title={LocaleStrings.AppLogoToolTip}
          />
          <div onClick={this.homeRedirect} className={styles.clbHeading} title={LocaleStrings.AppLogoToolTip}>
            {LocaleStrings.AppHeaderTitleLabel}
          </div>
        </Navbar.Brand>
        <div className={styles.navIconArea}>
          <div>
            <TooltipHost
              content={LocaleStrings.MoreInfoToolTip}
              delay={2}
              calloutProps={calloutProps}
              hostClassName={styles.cmpHeaderTooltipHostStyle}
            >
              <Icon className={styles.infoIcon} id={buttonId} iconName="Info" onClick={this.toggleIsCalloutVisible} />
            </TooltipHost>
            {this.state.isCalloutVisible && (
              <Callout
                className={styles.cmpHeaderInfoCallout}
                ariaLabelledBy={labelId}
                ariaDescribedBy={descriptionId}
                gapSpace={20}
                target={`#${buttonId}`}
                onDismiss={this.toggleIsCalloutVisible}
                setInitialFocus
                directionalHint={DirectionalHint.bottomCenter}
              >
                <Text block variant="xLarge" className={styles.infoCalloutTitle}>
                  {LocaleStrings.AboutHeaderLabel}
                </Text>
                <Text block variant="small" className={styles.infoCalloutTitleBody}>
                  {LocaleStrings.AboutContentLabel}
                </Text>
                <Text block variant="xLarge" className={styles.infoCalloutTitle}>
                  {LocaleStrings.AdditionalResourcesHeaderLabel}
                </Text>
                <Text block variant="small" className={styles.infoCalloutTitleBody}>
                  {LocaleStrings.AdditionalResourcesContentLabel}
                </Text>
                <Link href={Strings.M365Champions} target="_blank" className={`${styles.infoCalloutLink} ${styles.infoCalloutLinkFont}`}>
                  {LocaleStrings.M365ChampionCommunityLinkLabel}
                </Link>
                <Link href={Strings.DrivingAdoption} target="_blank" className={`${styles.infoCalloutLink} ${styles.infoCalloutLinkFont}`}>
                  {LocaleStrings.DrivingAdoptionLinkLabel}
                </Link>
                <Text block variant="xLarge" className={styles.infoCalloutTitle}>
                  ----
                </Text>
                <Text block variant="small">
                  {LocaleStrings.CurrentVersionLabel} {packageSolution.solution.version}
                </Text>
                <Text block variant="small">
                  {LocaleStrings.LatestVersionLabel} <Link href={Strings.LatestVersion} target="_blank">{LocaleStrings.CMPGitHubLinkLabel}</Link>
                </Text>
                <Text block variant="xLarge" className={styles.infoCalloutTitle}>
                  ----
                </Text>
                <Text block variant="small">
                  {LocaleStrings.VisitLabel}
                </Text>
                <Text block variant="small">
                  {LocaleStrings.OverviewLabel} <Link href={Strings.M365CMP} target="_blank">{LocaleStrings.MSAdoptionHubLinkLabel}</Link>
                </Text>
                <Text block variant="small">
                  {LocaleStrings.DocumentationLabel} <Link href={Strings.M365CmpApp} target="_blank">{LocaleStrings.CMPGitHubLinkLabel}</Link>
                </Text>
              </Callout>
            )}
          </div>
          <div>
            <a href={Strings.HelpUrl} target="_blank">
              <TooltipHost
                content={LocaleStrings.SupportToolTip}
                delay={2}
                calloutProps={calloutProps}
                hostClassName={styles.cmpHeaderTooltipHostStyle}
              >
                <Icon aria-label="Unknown" iconName="Unknown" className={styles.supportIcon} />
              </TooltipHost>
            </a>
          </div>
          <div>
            <a href={Strings.FeedbackUrl} target="_blank">
              <TooltipHost
                content={LocaleStrings.FeedbackToolTip}
                delay={2}
                calloutProps={calloutProps}
                hostClassName={styles.cmpHeaderTooltipHostStyle}
              >
                <Icon aria-label="Feedback" iconName="Feedback" className={styles.feedbackIcon} />
              </TooltipHost>
            </a>
          </div>
        </div>
      </Navbar>
    );
  }
}
