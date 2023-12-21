import { TooltipHost } from '@fluentui/react';
import "bootstrap/dist/css/bootstrap.min.css";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import {
  Popover, PopoverSurface,
  PopoverTrigger, Link, Text
} from '@fluentui/react-components';
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
  appTitle: string;
  currentThemeName?: string;
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
    const isDarkOrContrastTheme = this.props.currentThemeName === Strings.themeDarkMode || this.props.currentThemeName === Strings.themeContrastMode;
    return (
      <Navbar className={`${styles.navbg}${isDarkOrContrastTheme ? " " + styles.navbgDarkContrast : ""}`}>
        <Navbar.Brand href="#home" className={styles.white}>
          <img
            src={this.props.logoImageURL}
            className={`d-inline-block ${styles.clbHeaderLogo}`}
            alt="applogo"
            onClick={this.homeRedirect}
            title={LocaleStrings.AppLogoToolTip}
          />
          <div onClick={this.homeRedirect} className={styles.clbHeading} title={this.props.appTitle}>
            {this.props.appTitle}
          </div>
        </Navbar.Brand>
        <div className={styles.navIconArea}>
          <Popover
            withArrow={true}
            open={this.state.isCalloutVisible}
            inline={true}
            onOpenChange={this.toggleIsCalloutVisible}
            positioning="before"
            size='medium'
            closeOnScroll={true}
          >
            <PopoverTrigger disableButtonEnhancement={true}>
              <div
                onKeyDown={(evt: any) => { if (evt.key === Strings.stringEnter) this.toggleIsCalloutVisible() }}
                aria-label={LocaleStrings.MoreInfoToolTip}
                tabIndex={0}
                role="button"
                onClick={this.toggleIsCalloutVisible}
                className={styles.infoIconWrapper}
              >
                <TooltipHost
                  content={LocaleStrings.MoreInfoToolTip}
                  delay={2}
                  calloutProps={calloutProps}
                  hostClassName={styles.cmpHeaderTooltipHostStyle}
                >
                  <Icon className={styles.infoIcon} iconName="Info" />
                </TooltipHost>
              </div>
            </PopoverTrigger>
            <PopoverSurface as="div" className={styles.cmpHeaderInfoCallout}>
              <Text block as='h2' className={styles.infoCalloutTitle}>
                {LocaleStrings.AboutHeaderLabel} {this.props.appTitle}:
              </Text>
              <Text block as="p" className={styles.infoCalloutTitleBody}>
                {this.props.appTitle} {LocaleStrings.AboutContentLabel}
              </Text>
              <Text block as='h2' className={styles.infoCalloutTitle}>
                {LocaleStrings.AdditionalResourcesHeaderLabel}
              </Text>
              <Text block as="p" className={styles.infoCalloutTitleBody}>
                {LocaleStrings.AdditionalResourcesContentLabel}
              </Text>
              <Link href={Strings.M365Champions} target="_blank" className={`${styles.infoCalloutLink} ${styles.infoCalloutLinkFont}`}>
                {LocaleStrings.M365ChampionCommunityLinkLabel}
              </Link>
              <Link href={Strings.DrivingAdoption} target="_blank" className={`${styles.infoCalloutLink} ${styles.infoCalloutLinkFont}`}>
                {LocaleStrings.DrivingAdoptionLinkLabel}
              </Link>
              <Text block as="p" className={styles.infoCalloutTitle}>
                ----
              </Text>
              <Text block as="p">
                {LocaleStrings.CurrentVersionLabel} {packageSolution.solution.version}
              </Text>
              <Text block as="p">
                {LocaleStrings.LatestVersionLabel} <Link href={Strings.LatestVersion} target="_blank">{LocaleStrings.CMPGitHubLinkLabel}</Link>
              </Text>
              <Text block as="p" className={styles.infoCalloutTitle}>
                ----
              </Text>
              <Text block as="p">
                {LocaleStrings.VisitLabel}
              </Text>
              <Text block as="p">
                {LocaleStrings.OverviewLabel} <Link href={Strings.M365CMP} target="_blank">{LocaleStrings.MSAdoptionHubLinkLabel}</Link>
              </Text>
              <Text block as="p">
                {LocaleStrings.DocumentationLabel} <Link href={Strings.M365CmpApp} target="_blank">{LocaleStrings.CMPGitHubLinkLabel}</Link>
              </Text>
            </PopoverSurface>
          </Popover>
          <div>
            <a href={Strings.HelpUrl} target="_blank" aria-label="Support" role="link">
              <TooltipHost
                content={LocaleStrings.SupportToolTip}
                delay={2}
                calloutProps={calloutProps}
                hostClassName={styles.cmpHeaderTooltipHostStyle}
              >
                <Icon iconName="Unknown" className={styles.supportIcon} />
              </TooltipHost>
            </a>
          </div>
          <div>
            <a href={Strings.FeedbackUrl} target="_blank" aria-label="Feedback" role="link">
              <TooltipHost
                content={LocaleStrings.FeedbackToolTip}
                delay={2}
                calloutProps={calloutProps}
                hostClassName={styles.cmpHeaderTooltipHostStyle}
              >
                <Icon iconName="Feedback" className={styles.feedbackIcon} />
              </TooltipHost>
            </a>
          </div>
        </div>
      </Navbar >
    );
  }
}
