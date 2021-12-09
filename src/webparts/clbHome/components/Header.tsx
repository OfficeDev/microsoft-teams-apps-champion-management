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
            src={require("../assets/CMPImages/MsLogo.png")}
            className={`d-inline-block ${styles.clbHeaderLogo}`}
            alt="mslogo"
            onClick={this.homeRedirect}
            title="Home"
          />
          <span onClick={this.homeRedirect} className={styles.clbHeading} title="Home">Champion Management Platform</span>
        </Navbar.Brand>
        <Nav.Item className="ml-auto" style={{ marginRight: "1%" }}>
          <div className={styles.icon}>
            <TooltipHost
              content="More Info"
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
                  About the Champion Management Platform (CMP):
                </Text>
                <Text block variant="small" className={style.titlebody}>
                  Our Champion Management Platform was created with organizational Champions / Adoption Specialists in mind. Hearing from the Microsoft 365 Champion Community this app was developed to deliver a platform to help create and sustain your own communities. Starting with inspiration through execution in helping you achieve more within your own communities!
                </Text>
                <Text block variant="xLarge" className={style.title}>
                  Additional Resources:
                </Text>
                <Text block variant="small" className={style.titlebody}>
                  The Microsoft Teams Customer Advocacy Group is focused on delivering solutions like these to inspire and help you achieve your goals. Follow and join in through these other resources to learn more from us and the community:
                </Text>
                <Link href={Strings.M365Champions} target="_blank" className={`${style.link} ${style.linkFont}`}>
                  Microsoft 365 Champion Community
                </Link>
                <Link href={Strings.DrivingAdoption} target="_blank" className={`${style.link} ${style.linkFont}`}>
                  Driving Adoption on the Microsoft Technical Community
                </Link>
                <Text block variant="xLarge" className={style.title}>
                  ----
                </Text>
                <Text block variant="small">
                  Current Version: {packageSolution.solution.version}
                </Text>
                <Text block variant="small">
                  Latest Version: <Link href={Strings.LatestVersion} target="_blank">CMP GitHub</Link>
                </Text>
                <Text block variant="xLarge" className={style.title}>
                  ----
                </Text>
                <Text block variant="small">
                  Visit the Champion Management Platform pages to learn more:
                </Text>
                <Text block variant="small">
                  Overview & Information on our <Link href={Strings.M365CMP} target="_blank">Microsoft Adoption Hub</Link>
                </Text>
                <Text block variant="small">
                  Solution technical documentation and architectural overview on <Link href={Strings.M365CmpApp} target="_blank">GitHub</Link>
                </Text>
              </Callout>
            )}
          </div>

        </Nav.Item>
        <Nav.Item>
          <div className={styles.icon}>
            <a href={Strings.HelpUrl} target="_blank">
              <TooltipHost
                content="Support"
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
                content="Feedback"
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
