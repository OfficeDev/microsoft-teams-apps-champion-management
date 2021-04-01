import React, { Component } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import styles from "../scss/ClbHome.module.scss";
import Navbar from "react-bootstrap/Navbar";

interface IHeaderProps {
  showSearch: boolean;
  clickcallback: () => void; //will redirects to home
}
export default class Header extends Component<IHeaderProps, {}> {
  constructor(_props: any) {
    super(_props);
    this.homeRedirect = this.homeRedirect.bind(this);
  }
  public homeRedirect() {
    this.props.clickcallback();
  }
  public render() {
    return (
      <Navbar className={styles.navbg}>
        <Navbar.Brand href="#home" className={styles.white}>
          <img
            src={require("../assets/images/mslogo.png")}
            width="auto"
            height="40"
            className="d-inline-block"
            alt="mslogo"
            onClick={this.homeRedirect}
          />
          <span onClick={this.homeRedirect}>Champion Management Platform</span>
        </Navbar.Brand>
      </Navbar>
    );
  }
}
