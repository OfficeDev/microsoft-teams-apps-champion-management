import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { MSGraphClient } from "@microsoft/sp-http";
import {
  TeamsComponentContext,
  ConnectedComponent,
  Panel,
  PanelBody,
  PanelHeader,
  PanelFooter,
  Surface,
  ThemeStyle,
} from "msteams-ui-components-react";
import {
  getContext,
  primaryButton,
  anchor,
  compoundButton,
} from "msteams-ui-styles-core";
import {
  PrimaryButton,
  CompoundButton,
} from "office-ui-fabric-react/lib/Button";
import { initializeIcons } from "@uifabric/icons";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import {
  TeamsBaseComponent,
  ITeamsBaseComponentProps,
  ITeamsBaseComponentState,
} from "./TeamsBaseComponent";
import {
  MessageBar,
  MessageBarType,
} from "office-ui-fabric-react/lib/MessageBar";
import * as strings from "../constants/strings";
import "../assets/stylesheets/main.scss";
import * as $ from "jquery";
import IProfileImage from "../models/IProfileImage";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import dbStyles from "../scss/CMPDigitalBadge.module.scss";
import { Icon } from '@fluentui/react/lib/Icon';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import siteconfig from "../config/siteconfig.json";
import commonServices from '../Common/CommonServices';
import * as stringsConstants from "../constants/strings";
import { List } from '@fluentui/react/lib/List';
import { Label } from "@microsoft/office-ui-fabric-react-bundle";


const config = {
  baseFontSize: 16,
  style: ThemeStyle.Light,
};
const contextCSS = getContext(config);
const graphUrl = "https://graph.microsoft.com";
const graphMyPhotoApiUrl = graphUrl + "/v1.0/me/photo";
const graphMyPhotoBitsUrl = graphMyPhotoApiUrl + "/$value";
let upn: string | undefined = "";

//CSS classes for Image List
const classNames = mergeStyleSets({
  listGrid: {
    width: '100%',
    overflow: 'hidden',
    fontSize: 0,
    position: 'relative',
    margin: 'auto auto auto auto'
  },
  listGridTile: {
    width: '25%',
    marginBottom: '8%',
    textAlign: 'center',
    outline: 'none',
    position: 'relative',
    float: 'left',
    selectors: {
      'focus:after': {
        content: '',
        position: 'absolute',
        left: 2,
        right: 2,
        top: 2,
        bottom: 2,
        boxSizing: 'border-box',
      },
    },
  },
  listGridSizer: {
    paddingBottom: '100%',
  },
  listGridPadder: {
    position: 'absolute',
    align: 'center',
    left: 2,
    top: 2,
    right: 2,
    bottom: 2,
    marginRight: '10%',
    marginBottom: '5%'
  },
  listGridLabel: {
    background: '#464775',
    color: '#FFFFFF',
    position: 'absolute',
    bottom: 0,
    left: 0,
    width: '100%',
    height: '28px',
    boxSizing: 'border-box',
    font: "normal normal bold 16px/21px Segoe UI",
    paddingTop: '3px',
    overflowX: 'hidden'
  },
  listGridImage: {
    position: 'absolute',
    top: '15%',
    left: 0,
    width: '70%',
    marginLeft: "15%"
  },
  acceptIcon: {
    marginRight: "10px",
    fontSize: "16px",
    fontWeight: "bolder",
    color: "#FFFFFF",
    opacity: 1
  },
  downloadIcon: {
    fontSize: "16px",
    fontWeight: "bolder",
    opacity: 1
  }
});

export interface IDigitalBadgeState extends ITeamsBaseComponentState {
  entityId?: string;
  isLoading: boolean;
  themeLoaded: boolean;
  profileImage?: IProfileImage;
  isLoggedIn: boolean;
  hasAccepted: boolean;
  hasImageSelected: boolean;
  imageURL: string;
  isApplying: boolean;
  isApplied: boolean;
  error: string;
  upn?: string;
  imageDownloaded: boolean;
  downloadText: string;
  showAccept: boolean;
  siteUrl: string;
  userletters: string;
  sitename: string;
  inclusionpath: string;
  allBadgeImages: string[];
  noBadgesFlag: boolean;
}
export interface IDigitalBadgeProps extends ITeamsBaseComponentProps {
  clientId: string;
  description: string;
  context: WebPartContext;
  clickcallback: () => void;
  clickcallchampionview: () => void;
  siteUrl: string;
}

export default class DigitalBadge extends TeamsBaseComponent<
  IDigitalBadgeProps,
  IDigitalBadgeState
> {
  constructor(props: IDigitalBadgeProps, states: IDigitalBadgeState) {
    super(props, states);
    this._onDownloadImage = this._onDownloadImage.bind(this);
    this.onUserAcceptance = this.onUserAcceptance.bind(this);
    this.onBadgeSelected = this.onBadgeSelected.bind(this);
    this._onApplyProfileImage = this._onApplyProfileImage.bind(this);
    this.getPhotoBits = this.getPhotoBits.bind(this);
    this.getAllBadgeImages = this.getAllBadgeImages.bind(this);
    this.onRenderCell = this.onRenderCell.bind(this);
  }
  private _requestOptions: {} = {
    headers: {
      "X-ClientTag": "NONISV|Microsoft|ChampBadge/1.0.0",
    },
  };
  public componentWillMount() {
    initializeIcons();
    let profile: IProfileImage = { url: "", width: 0 };

    this.setState({
      fontSize: this.pageFontSize(),
      isLoading: true,
      themeLoaded: false,
      profileImage: profile,
      hasAccepted: false,
      hasImageSelected: false,
      imageURL: "",
      isApplying: false,
      isApplied: false,
      isLoggedIn: false,
      error: "",
      imageDownloaded: false,
      showAccept: false,
      downloadText: strings.DownloadButtonText,
      userletters: "",
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      siteUrl: this.props.siteUrl,
      allBadgeImages: [],
      noBadgesFlag: false
    });

    this.forceUpdate();
    setTimeout(() => {
      this._renderListAsync();
    }, 100);

  }

  //Multiple Badges   
  //Render Fluent UI list cell to show the images and hyperlinks
  private onRenderCell(item: any, index: number | undefined) {
    try {
      return (
        <div
          className={classNames.listGridTile}
          data-is-focusable
        >
          <div className={classNames.listGridSizer}>
            <div className={classNames.listGridPadder}>
              <a href="#" onClick={() => { this.onBadgeSelected(item.url); }}>
                <img src={item.url} className={classNames.listGridImage} /></a>
              <span className={classNames.listGridLabel} title={item.title}>{item.title}</span>
            </div>
          </div>
        </div>
      );
    }
    catch (error) {
      console.error("CMP_DigitalBadge_onRenderCell \n", JSON.stringify(error));
      this.setState({
        error: stringsConstants.TOTErrorMessage + "while displaying the digital badges. Below are the details: \n" + JSON.stringify(error),
        isApplying: false,
        isLoading: false
      });
    }

  }

  //On a badge selection set the state and show the overlay images
  public onBadgeSelected(img: string): void {
    try {
      this.setState({
        hasImageSelected: true,
        imageURL: img,
      });
      this.showUserInformation(upn);
    }
    catch (error) {
      console.error("CMP_DigitalBadge_onBadgeSelected \n", JSON.stringify(error));
      this.setState({
        error: stringsConstants.TOTErrorMessage + "while applying the digital badge. Below are the details: \n" + JSON.stringify(error),
        isApplying: false,
        isLoading: false
      });
    }
  }

  //Get badge images from 'Digital Badge Assets' document library.
  private async getAllBadgeImages() {
    try {
      this.setState({ isLoading: true });
      let commonServiceManager: commonServices = new commonServices(this.props.context, this.props.siteUrl);
      const resultImages: any[] = await commonServiceManager.getAllBadgeImages(stringsConstants.DigitalBadgeLibrary, this.props.context.pageContext.user.email.toLowerCase());
      if (resultImages.length == 0)
        this.setState({ noBadgesFlag: true, isLoading: false });
      else
        this.setState({ allBadgeImages: resultImages, isLoading: false });
    }
    catch (error) {
      console.error("CMP_DigitalBadge_getAllBadgeImages \n", JSON.stringify(error));
      this.setState({
        error: stringsConstants.TOTErrorMessage + "while retrieving the digital badges. Below are the details: \n" + JSON.stringify(error),
        isApplying: false,
        isLoading: false
      });
    }
  }
  //Multiple Badges Section Ends

  //Get currents user's details from Mmber List for Digital Badge processing
  private _renderListAsync() {
    microsoftTeams.initialize();
    microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
    microsoftTeams.getContext((context: microsoftTeams.Context) => {
      this.props.context.spHttpClient
        .get(
          "/" + this.state.inclusionpath + "/" + this.state.sitename +
          "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
          SPHttpClient.configurations.v1
        )
        .then((responseuser: SPHttpClientResponse) => {
          responseuser.json().then((datauser: any) => {
            this.props.context.spHttpClient
              .get(
                "/" + this.state.inclusionpath + "/" + this.state.sitename +
                "/_api/web/lists/GetByTitle('Member List')/Items?$filter=Title eq '" + datauser.Email.toLowerCase() + "'",
                SPHttpClient.configurations.v1
              )
              .then((response: SPHttpClientResponse) => {
                response.json().then((datada) => {
                  let dataexists = datada.value.find(
                    (x: any) =>
                      x.Title.toLowerCase() == datauser.Email.toLowerCase()
                  );
                  if (dataexists) {
                    let userassignletters = "";
                    let usernamearray = datauser.DisplayName.split(" ");
                    if (usernamearray.length === 1) {
                      userassignletters = usernamearray[0][0].toUpperCase();
                    } else if (usernamearray.length > 1) {
                      userassignletters =
                        usernamearray[0][0].toUpperCase() +
                        usernamearray[
                          usernamearray.length - 1
                        ][0].toUpperCase();
                    }
                    this.setState({
                      showAccept: true,
                      userletters: userassignletters,
                    });
                  }
                  this.updateTheme(context.theme);
                  upn = datauser.Email;
                  this.setState({
                    isLoading: false,
                    entityId: context.entityId,
                    upn: context.upn,
                  });
                  this.showUserInformation(upn);
                });
              });
          });
        });
    });
  }


  public render(): React.ReactElement<IDigitalBadgeProps> {
    return (
      <div>
        {this.state.isLoading && (
          <Spinner
            size={SpinnerSize.large}
            ariaLabel={strings.LoadingSpinnerLabel}
            label={strings.LoadingSpinnerLabel}
            ariaLive="assertive"
          />
        )}
        {!this.state.isLoading && (
          <TeamsComponentContext
            fontSize={this.state.fontSize}
            theme={this.state.theme}
          >
            <ConnectedComponent
              render={(props: { context: any }) => {
                const { context } = props;
                const { rem, font, colors, style } = context;
                const { sizes, weights } = font;
                contextCSS.style = this.state.theme;
                const styleProps: any = {};
                switch (style) {
                  case ThemeStyle.Dark:
                    styleProps.color = colors.dark.brand00;
                    break;
                  case ThemeStyle.HighContrast:
                    styleProps.color = colors.highContrast.white;
                    break;
                  case ThemeStyle.Light:
                  default:
                    styleProps.color = colors.light.brand00;
                }
                const styles = {
                  header: { ...sizes.title, ...weights.semibold },
                  section: {
                    ...sizes.base,
                    marginTop: rem(1.4),
                    marginBottom: rem(1.4),
                  },
                  footer: { ...sizes.xsmall },
                  div: {},
                };
                const anchorClass: string = anchor(contextCSS);

                return (
                  <Surface>
                    <Panel>
                      <PanelHeader>
                        <div className={dbStyles.digitalBadgePath}>
                          <img src={require("../assets/CMPImages/BackIcon.png")}
                            className={dbStyles.backImg}
                          />
                          <span
                            className={dbStyles.backLabel}
                            onClick={this.props.clickcallback}
                            title="Back"
                          >
                            Back
                          </span>
                          <span className={dbStyles.border}></span>
                          <span className={dbStyles.digitalBadgeLabel}>Digital Badge</span>
                        </div>
                      </PanelHeader>
                      <PanelBody>
                        <div className={"DigitalBadge"} style={styles.section}>
                          <div className={`container`}>
                            {this.state.isLoading && (
                              <Spinner
                                size={SpinnerSize.large}
                                ariaLabel={strings.LoadingSpinnerLabel}
                                label={strings.LoadingSpinnerLabel}
                                ariaLive="assertive"
                              />
                            )}
                            {!this.state.isLoading && (
                              <section
                                aria-live="polite"
                                className={"contentSection"}
                              >
                                <div
                                  className={dbStyles.introPageBox}
                                >
                                  {!this.state.hasAccepted && (
                                    <div className={dbStyles.divChild1}>
                                      <span className={dbStyles.imgText}>
                                        {strings.PreAcceptPageTitle}
                                      </span>
                                      <br /> <br />
                                      <img
                                        src={require("../assets/CMPImages/AppBanner.png")}
                                        className={"bannerimage"}
                                        alt={strings.BannerImageAlt}
                                      />
                                      {this.state.badgeImgURL}
                                    </div>
                                  )}
                                  <div className={dbStyles.divChild2}>
                                    {!this.state.hasAccepted &&
                                      this.state.showAccept && (
                                        <p
                                          className={"description"}
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.PreAcceptDisclaimer
                                          )}
                                        />
                                      )}
                                    <br />
                                    {!this.state.hasAccepted &&
                                      this.state.showAccept && (
                                        <p
                                          className={"description"}
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.PreAcceptDisclaimer2
                                          )}
                                        />
                                      )}
                                    {!this.state.hasAccepted &&
                                      !this.state.showAccept && (
                                        <div>
                                          <p
                                            className={"description"}
                                            dangerouslySetInnerHTML={this.createMarkup(
                                              strings.NotQualifiedPreAcceptDisclaimer
                                            )}
                                          />
                                          <p onClick={this.props.clickcallback}>
                                            How to get Champion Badge
                                          </p>
                                        </div>
                                      )}
                                  </div>
                                </div>
                                {this.state.hasAccepted && !this.state.hasImageSelected && (
                                  <div className={dbStyles.badgeList}>
                                    <p
                                      dangerouslySetInnerHTML={this.createMarkup(
                                        strings.MultipleBadgeMessage
                                      )}
                                    />
                                    {this.state.noBadgesFlag && (
                                      <p
                                        className={"description"}
                                        dangerouslySetInnerHTML={this.createMarkup(
                                          strings.NoBadgeMessage
                                        )}
                                      />
                                    )}
                                    <List
                                      className={classNames.listGrid}
                                      items={this.state.allBadgeImages}
                                      renderedWindowsAhead={6}
                                      onRenderCell={this.onRenderCell}
                                    />
                                  </div>
                                )}
                                {this.state.hasAccepted && this.state.hasImageSelected && (
                                  <div className={dbStyles.badgeDetailsHeading}>
                                    {strings.PageTitle}
                                  </div>
                                )}
                                <div className={dbStyles.badgeDetailsContainer}>
                                  <div className={dbStyles.badgeBtnArea}>
                                    <div className={`profileContainer ${dbStyles.profileArea}`}>
                                      {this.state.profileImage.url &&
                                        this.state.hasAccepted &&
                                        this.state.hasImageSelected &&
                                        this.state.profileImage.url !==
                                        "../assets/images/noimage.png" && (
                                          <div
                                            id="forDomToImage"
                                            style={{ maxWidth: "700px" }}
                                          >
                                            <img
                                              style={{
                                                width: `150px`,
                                              }}
                                              src={this.state.profileImage.url}
                                              id={"profileImage"}
                                              alt={strings.ProfileImageAlt}
                                            />
                                            <img
                                              style={{
                                                width: `150px`,
                                                marginTop: `-150px`,
                                              }}
                                              id={"badgeImage"}
                                              alt={strings.BadgeImageAlt}
                                              src={this.state.imageURL}
                                            />
                                          </div>
                                        )}
                                      {this.state.profileImage.url &&
                                        this.state.profileImage.url ===
                                        "../assets/images/noimage.png" &&
                                        this.state.hasAccepted &&
                                        this.state.hasImageSelected && (
                                          <div
                                            id="forDomToImage"
                                            style={{ maxWidth: "700px" }}
                                          >
                                            <img
                                              src={require("../assets/images/noimage.png")}
                                              style={{ width: `100px` }}
                                              id={"profileImage"}
                                              alt={strings.ProfileImageAlt}
                                            />
                                            <div className={"profiletext"}>
                                              {this.state.userletters}
                                            </div>
                                            <img
                                              style={{
                                                width: `100px`,
                                                marginTop: `-100px`,
                                              }}
                                              id={"badgeImage"}
                                              alt={strings.BadgeImageAlt}
                                              src={this.state.imageURL}
                                            />
                                          </div>
                                        )}
                                      {!this.state.profileImage.url &&
                                        this.state.hasAccepted &&
                                        this.state.hasImageSelected && (
                                          <div>
                                            <img
                                              src={require("../assets/images/noprofile.png")}
                                              id={"photoStuff"}
                                              alt={"strings.NoProfileImageAlt"}
                                              aria-hidden="true"
                                              style={{ width: "150px", height: "auto" }}
                                            />
                                          </div>
                                        )}
                                    </div>
                                    {!this.state.isApplying &&
                                      this.state.profileImage.url &&
                                      this.state.hasAccepted &&
                                      this.state.hasImageSelected &&
                                      !this.state.isApplying &&
                                      !this.state.isApplied && (
                                        <div className={`buttonContainer ${dbStyles.buttonArea}`}>
                                          <PrimaryButton
                                            className={primaryButton(
                                              contextCSS
                                            )}
                                            onClick={this._onApplyProfileImage}
                                            ariaLabel={strings.ApplyButtonText}
                                            ariaDescription={
                                              strings.ApplyButtonAriaDescription
                                            }
                                            disabled={
                                              this.state.isApplying ||
                                              this.state.isApplied ||
                                              this.state.error.length > 0
                                            }
                                            title="Apply"
                                          >
                                            <Icon iconName="Completed" className={`${classNames.acceptIcon}`} />
                                            {strings.ApplyButtonText}
                                          </PrimaryButton>
                                          <br />
                                          {this.state.profileImage.url !==
                                            "../assets/images/noimage.png" && (
                                              <div className={dbStyles.downloadArea}>
                                                <CompoundButton
                                                  iconProps={{
                                                    iconName: "Download",
                                                    className: classNames.downloadIcon
                                                  }}
                                                  className={`
                                                    ${compoundButton(contextCSS)
                                                      .container} ${dbStyles.downloadBtn}`
                                                  }
                                                  style={{ ...styleProps }}
                                                  title="Download"
                                                  onClick={this._onDownloadImage}
                                                  ariaLabel={
                                                    strings.DownloadButtonText
                                                  }
                                                  ariaDescription={
                                                    strings.DownloadButtonAriaDescription
                                                  }
                                                  disabled={
                                                    this.state.isApplying ||
                                                    this.state.isApplied ||
                                                    this.state.imageDownloaded
                                                  }
                                                >
                                                  {this.state.downloadText}
                                                </CompoundButton> <br />
                                                <Label
                                                  style={styleProps}
                                                  className={dbStyles.downloadMsg}
                                                >
                                                  {
                                                    this.state.imageDownloaded
                                                      ? strings.DownloadedButtonSecondaryText
                                                      : strings.DownloadButtonSecondaryText

                                                  }
                                                </Label>
                                              </div>
                                            )}
                                        </div>
                                      )}
                                  </div>
                                  <div className={dbStyles.badgeDetailsText}>
                                    {this.state.hasAccepted &&
                                      this.state.hasImageSelected &&
                                      !this.state.isApplied &&
                                      this.state.profileImage.url &&
                                      this.state.profileImage.url !==
                                      "../assets/images/noimage.png" && (
                                        <p
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.PreApplyDisclaimer,
                                            anchorClass
                                          )}
                                        />
                                      )}
                                    {this.state.hasAccepted &&
                                      this.state.hasImageSelected &&
                                      !this.state.isApplied &&
                                      this.state.profileImage.url && (
                                        <p
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.PreApplyDisclaimer1,
                                            anchorClass
                                          )}
                                        />
                                      )}
                                  </div>
                                </div>
                                {!this.state.profileImage.url &&
                                  this.state.hasAccepted &&
                                  this.state.hasImageSelected && (
                                    <p
                                      dangerouslySetInnerHTML={this.createMarkup(
                                        strings.NoProfileImageDescription,
                                        anchorClass
                                      )}
                                    />
                                  )}
                                {!this.state.isApplying && (
                                  <div className={"buttonContainer"}>
                                    {!this.state.hasAccepted &&
                                      this.state.showAccept && (
                                        <PrimaryButton
                                          className={primaryButton(contextCSS)}
                                          onClick={this.onUserAcceptance}
                                          ariaLabel={strings.AcceptButtonText}
                                          ariaDescription={
                                            strings.AcceptButtonAriaDescription
                                          }
                                          title="Accept"
                                        >
                                          <Icon iconName="Completed" className={`${classNames.acceptIcon}`} />
                                          {strings.AcceptButtonText}
                                        </PrimaryButton>
                                      )}
                                    {!this.state.hasAccepted &&
                                      !this.state.showAccept && (
                                        <p
                                          className={"description"}
                                          style={{ color: "red" }}
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.UnauthorizedText
                                          )}
                                        />
                                      )}
                                  </div>
                                )}
                                {this.state.isApplying &&
                                  !this.state.isApplied && (
                                    <div className={"applySpinnerContainer"}>
                                      <Spinner
                                        ariaLabel={strings.ApplySpinnerLabel}
                                        size={SpinnerSize.large}
                                        label={strings.ApplySpinnerLabel}
                                        ariaLive="assertive"
                                      />
                                    </div>
                                  )}
                                {this.state.isApplied &&
                                  !this.state.isApplying && (
                                    <div className={"messagingContainer"}>
                                      <MessageBar
                                        ariaLabel={strings.SuccessMessage}
                                        messageBarType={MessageBarType.success}
                                      >
                                        <span
                                          dangerouslySetInnerHTML={this.createMarkup(
                                            strings.SuccessMessage,
                                            anchorClass
                                          )}
                                        />
                                      </MessageBar>
                                    </div>
                                  )}
                                {this.state.error && (
                                  <div className={"messagingContainer"}>
                                    <MessageBar
                                      ariaLabel={this.state.error}
                                      messageBarType={MessageBarType.error}
                                    >
                                      {this.state.error}
                                    </MessageBar>
                                  </div>
                                )}
                              </section>
                            )}
                            <canvas
                              id="profileCanvas"
                              width={this.state.profileImage.width}
                              height={this.state.profileImage.width}
                              style={{ display: "none" }}
                            ></canvas>
                            <canvas
                              id="profileCanvasDownload"
                              width={this.state.profileImage.width}
                              height={this.state.profileImage.width}
                              style={{ display: "none" }}
                            ></canvas>
                          </div>
                        </div>
                        <div hidden={true} style={styles.section}>
                          {this.state.entityId}
                        </div>
                      </PanelBody>
                      <PanelFooter>
                        <div style={styles.footer}></div>
                      </PanelFooter>
                    </Panel>
                  </Surface>
                );
              }}
            ></ConnectedComponent>
          </TeamsComponentContext>
        )}
      </div>
    );
  }
  public createMarkup(markup: string, anchorClass: string = "") {
    if (markup.indexOf(strings.ANCHOR_ID) !== -1 && anchorClass !== "") {
      markup = markup.replace(strings.ANCHOR_ID, anchorClass);
    }
    return { __html: markup };
  }
  public getPhotoBits(): Promise<any> {
    let canvas: any = document.getElementById("profileCanvas");
    if (canvas.msToBlob) {
      // for IE
      console.log("Function msToBlob found. Using existing function.");
      return new Promise<Blob>((resolve: (arg0: any) => void) => {
        resolve(canvas.msToBlob());
      });
    } else {
      // other browsers ** this isn't currently working **
      console.log("Function msToBlob not found. Using custom function.");
      return this._getCanvasBlob(canvas)
        .then((blob: any) => {
          return blob;
        })
        .catch((errDb: string) => {
          console.error("getPhotoBits error: " + errDb);
          this.setState({
            error: "error",
            isApplying: false,
          });
        });
    }
  }

  private _onRenderCanvas(profileImage: IProfileImage): Promise<any> {
    let promise: Promise<any> = new Promise<any>(
      (resolve: any, _reject: any) => {
        const canvas: any = document.getElementById("profileCanvas");
        const canvasDownload: any = document.getElementById(
          "profileCanvasDownload"
        );
        const context = canvas.getContext("2d");
        const contextDownload = canvasDownload.getContext("2d");
        const profileImageObj: HTMLImageElement = new Image();
        const badgeImageObj: HTMLImageElement = new Image();
        profileImageObj.src = profileImage.url;
        badgeImageObj.src = this.state.imageURL;
        profileImageObj.onload = () => {
          context.drawImage(
            profileImageObj,
            0,
            0,
            `${profileImage.width}`,
            `${profileImage.width}`
          );
          contextDownload.drawImage(
            profileImageObj,
            0,
            0,
            `${profileImage.width}`,
            `${profileImage.width}`
          );
          context.drawImage(
            badgeImageObj,
            0,
            0,
            `${profileImage.width}`,
            `${profileImage.width}`
          );
        };
        resolve(profileImage);
      }
    );
    return promise;
  }

  private _onRenderCanvasNoImage(profileImage: IProfileImage): Promise<any> {
    let promise: Promise<any> = new Promise<any>(
      (resolve: any, _reject: any) => {
        const canvas: any = document.getElementById("profileCanvas");
        const canvasDownload: any = document.getElementById(
          "profileCanvasDownload"
        );
        const context = canvas.getContext("2d");
        const contextDownload = canvasDownload.getContext("2d");
        const profileImageObj: HTMLImageElement = new Image();
        const badgeImageObj: HTMLImageElement = new Image();
        profileImage.width = 100;
        profileImage.url = "../assets/images/noimage.png";
        this.setState({
          profileImage: profileImage,
        });
        profileImageObj.src = require("../assets/images/noimage.png");
        badgeImageObj.src = this.state.imageURL;
        profileImageObj.onload = () => {
          context.font = "32px Arial";
          context.textAlign = "center";
          context.fillText(
            this.state.userletters,
            canvas.width / 2,
            canvas.height / 2
          );
          context.drawImage(profileImageObj, 100, 20, 100, 100);
          contextDownload.drawImage(profileImageObj, 0, 0);
          context.drawImage(
            badgeImageObj,
            0,
            0,
            `${profileImage.width}`,
            `${profileImage.width}`
          );
        };
        resolve(profileImage);
      }
    );
    return promise;
  }

  private _onApplyProfileImage() {
    this.setState({
      isApplying: true,
    });

    this.getPhotoBits()
      .then((res: any) => {
        this.updateUserPhoto(res)
          .then((response: { status: number }) => {
            console.log(response);
            if (response.status === 200) {
              this.setState({
                isApplying: false,
                isApplied: true,
              });
            } else {
              this.setState({
                error: strings.ErrorMessage,
                isApplying: false,
              });
            }
          })
          .catch((errDb: string) => {
            console.error("_updateUserPhoto error: " + errDb);
            this.setState({
              error: strings.ErrorMessage,
              isApplying: false,
            });
          });
      })
      .catch((errDb: string) => {
        console.error("_updatePhoto error: " + errDb);
      });
  }

  private _getCanvasBlob(canvas: {
    toBlob: (arg0: (blob: any) => void) => void;
  }) {
    return new Promise<Blob>((resolve: (arg0: any) => void, _reject: any) => {
      canvas.toBlob((blob: any) => {
        resolve(blob);
      });
    });
  }

  private _onDownloadImage() {
    this.setState({
      imageDownloaded: true,
      downloadText: strings.DownloadingButtonText,
    });
    let canvasDownload: any = document.getElementById("profileCanvasDownload");
    let link: HTMLAnchorElement = document.createElement("a");

    if (canvasDownload.msToBlob) {
      // for IE
      let blob = canvasDownload.msToBlob();
      this.setState({ downloadText: strings.DownloadedButtonText });
    } else {
      // other browsers
      canvasDownload.toBlob((blob: any) => {
        let url = URL.createObjectURL(blob);
        link.href = url;

        link.setAttribute("download", "myProfileImage.jpg");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        this.setState({ downloadText: strings.DownloadedButtonText });
      });
    }
  }

  //On user acceptance show the digital badges for the user to select
  public onUserAcceptance(): void {
    this.setState({
      hasAccepted: true,
    });
    this.getAllBadgeImages();
  }

  public updateUserPhoto(blob: any): Promise<any> {
    let photoPromise: Promise<any> = new Promise(
      (resolve: (arg0: Response) => void, _reject: any) => {
        blob.lastModifiedDate = new Date();
        blob.name = "profile.jpeg";
        this.props.context.msGraphClientFactory
          .getClient()
          .then((client: MSGraphClient) => {
            client
              .api("me/photo/$value")
              .version("v1.0")
              .header("Content-Type", "image/jpeg")
              .responseType("json")
              .put(blob, (errDb, _res, rawresponse) => {
                if (!errDb) {
                  resolve(rawresponse);
                }
              });
          });
      }
    );
    return photoPromise;
  }

  public IsValidJSONString = (str: any) => {
    try {
      JSON.parse(str);
    } catch (e) {
      return false;
    }
    return true;
  }

  public getgraphMyPhotoBitsUrl(): Promise<any> {
    let photoPromise: Promise<any> = new Promise(
      (resolve: (arg0: Response) => void, reject: any) => {
        this.props.context.msGraphClientFactory
          .getClient()
          .then((garphClient: MSGraphClient) => {
            garphClient
              .api(graphMyPhotoBitsUrl)
              .version("v1.0")
              .headers({ "Content-Type": "blob", responseType: "blob" })
              .responseType("blob")
              .get()
              .then((data) => {
                resolve(data);
              })
              .catch((errDb) => {
                reject(errDb);
              });
          });
      }
    );
    return photoPromise;
  }

  public getgraphMyPhotoApiUrl(): Promise<any> {
    let photoPromise: Promise<any> = new Promise(
      (resolve: (arg0: Response) => void, _reject: any) => {
        this.props.context.msGraphClientFactory
          .getClient()
          .then((garphClient: MSGraphClient) => {
            garphClient
              .api(graphMyPhotoApiUrl)
              .version("v1.0")
              .headers({ "Content-Type": "blob", responseType: "blob" })
              .responseType("json")
              .get()
              .then((data) => {
                resolve(data);
              });
          });
      }
    );
    return photoPromise;
  }

  public showUserInformation(_upn1: string) {
    let currentProfileImageObj: IProfileImage = { url: "", width: 0 };
    this.getgraphMyPhotoBitsUrl()
      .then((blob) => {
        let blobUrl = URL.createObjectURL(blob);
        currentProfileImageObj.url = blobUrl;
        $("#photoStuff").attr("src", blobUrl);
        currentProfileImageObj.url = blobUrl;
      })
      .then((_asd) => {
        this.getgraphMyPhotoApiUrl().then((json) => {
          currentProfileImageObj.width = json.width;
          this.setState({
            profileImage: currentProfileImageObj,
          });
          this.forceUpdate();
          this._onRenderCanvas(currentProfileImageObj);
        });
      })
      .catch((error) => {
        if (error.statusCode === 404) {
          this._onRenderCanvasNoImage(currentProfileImageObj);
        }
      });
  }

  public getQueryParameters(): { upn: string; tenantId: string } {
    let queryParams = { upn: "", tenantId: "" };
    alert("Request query string params: " + location.search);
    location.search
      .substr(1)
      .split("&")
      .forEach((item) => {
        let s = item.split("="),
          k = s[0],
          v = s[1] && decodeURIComponent(s[1]);
        queryParams["upn"] = k;
        queryParams["tenantId"] = v;
      });
    return queryParams;
  }
}
