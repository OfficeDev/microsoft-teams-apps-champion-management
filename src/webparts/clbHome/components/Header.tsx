import React, { Component } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "../scss/CMPHome.module.scss";
import Nav from "react-bootstrap/Nav";
import Navbar from "react-bootstrap/Navbar";
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost, ITooltipHostStyles } from '@fluentui/react/lib/Tooltip';
import { Callout, Link, Text, FontWeights } from 'office-ui-fabric-react';
import * as Strings from '../constants/strings';
import * as LocaleStrings from 'ClbHomeWebPartStrings';

// Import package version
const packageSolution: any = require("../../../../config/package-solution.json");
const calloutProps = { gapSpace: 0 };

const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block', cursor: 'pointer' } };

const classes = mergeStyleSets({
  Icon: {
    fontSize: '20px',
    color: '#FFFFFF',
    opacity: 1,
    cursor: 'pointer'
  },
  feedbackIcon: {
    fontSize: '20px',
    color: '#FFFFFF',
    opacity: 1,
    cursor: 'pointer'
  }
});
const style = mergeStyleSets({
  button: {
    width: 130,
  },
  callout: {
    width: 620,
    padding: '10px 24px 20px 24px'
  },
  title: {
    marginBottom: 12,
    fontWeight: FontWeights.bold,
  },
  titlebody: {
    marginBottom: 12,
  },
  titlelink: {
    fontWeight: FontWeights.bold,
    color: "#6264A7"
  },
  link: {
    display: 'block',
    marginBottom: 12,
  },
  linkFont: {
    fontSize: "16px"
  }
});

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
          <span onClick={this.homeRedirect} className={styles.clbHeading} title={LocaleStrings.AppLogoToolTip}>{LocaleStrings.AppHeaderTitleLabel}</span>
        </Navbar.Brand>
        <Nav.Item className="ml-auto" style={{ marginRight: "1%" }}>
          <div className={styles.icon}>
            <TooltipHost
              content={LocaleStrings.MoreInfoToolTip}
              delay={2}
              calloutProps={calloutProps}
              styles={hostStyles}
            >
              <Icon className={classes.Icon} id={buttonId} iconName="Info" onClick={this.toggleIsCalloutVisible} />
            </TooltipHost>
            {this.state.isCalloutVisible && (
              <Callout
                className={style.callout}
                ariaLabelledBy={labelId}
                ariaDescribedBy={descriptionId}
                gapSpace={20}
                target={`#${buttonId}`}
                onDismiss={this.toggleIsCalloutVisible}
                setInitialFocus
                directionalHint={3}
              >
                <Text block variant="xLarge" className={style.title}>
                {LocaleStrings.AboutHeaderLabel}
                </Text>
                <Text block variant="small" className={style.titlebody}>
                {LocaleStrings.AboutContentLabel}
                </Text>
                <Text block variant="xLarge" className={style.title}>
                {LocaleStrings.AdditionalResourcesHeaderLabel}
                </Text>
                <Text block variant="small" className={style.titlebody}>
                {LocaleStrings.AdditionalResourcesContentLabel}
                </Text>
                <Link href={Strings.M365Champions} target="_blank" className={`${style.link} ${style.linkFont}`}>
                {LocaleStrings.M365ChampionCommunityLinkLabel}
                </Link>
                <Link href={Strings.DrivingAdoption} target="_blank" className={`${style.link} ${style.linkFont}`}>
                {LocaleStrings.DrivingAdoptionLinkLabel}
                </Link>
                <Text block variant="xLarge" className={style.title}>
                  ----
                </Text>
                <Text block variant="small">
                {LocaleStrings.CurrentVersionLabel} {packageSolution.solution.version}
                </Text>
                <Text block variant="small">
                {LocaleStrings.LatestVersionLabel} <Link href={Strings.LatestVersion} target="_blank">{LocaleStrings.CMPGitHubLinkLabel}</Link>
                </Text>
                <Text block variant="xLarge" className={style.title}>
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

        </Nav.Item>
        <Nav.Item>
          <div className={styles.icon}>
            <a href={Strings.HelpUrl} target="_blank">
              <TooltipHost
                content={LocaleStrings.SupportToolTip}
                delay={2}
                calloutProps={calloutProps}
                styles={hostStyles}
              >
                <Icon aria-label="Unknown" iconName="Unknown" className={classes.Icon} />
              </TooltipHost>
            </a>
          </div>
        </Nav.Item>
        <Nav.Item>
          <div className={styles.fbIcon}>
            <a href={Strings.FeedbackUrl} target="_blank">
              <TooltipHost
                content={LocaleStrings.FeedbackToolTip}
                delay={2}
                calloutProps={calloutProps}
                styles={hostStyles}
              >
                <Icon aria-label="Feedback" iconName="Feedback" className={classes.feedbackIcon} />
              </TooltipHost>
            </a>
          </div>
        </Nav.Item>
      </Navbar>
    );
  }
}
