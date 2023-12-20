import * as React from "react";
import {
  Icon, List, IRectangle, PrimaryButton, MessageBar,
  MessageBarType, Spinner, SpinnerSize, DefaultButton
} from '@fluentui/react';
import { MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import commonServices from '../Common/CommonServices';
import siteconfig from "../config/siteconfig.json";
import * as strings from "../constants/strings";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import dbStyles from "../scss/CMPDigitalBadge.module.scss";

const graphUrl = "https://graph.microsoft.com";
const graphMyPhotoApiUrl = graphUrl + "/v1.0/me/photo";
const graphMyPhotoBitsUrl = graphMyPhotoApiUrl + "/$value";
let upn: string | undefined = "";

export interface IDigitalBadgeState {
  isLoading: boolean;
  themeLoaded: boolean;
  profileImage: IProfileImage;
  isLoggedIn: boolean;
  hasAccepted: boolean;
  hasImageSelected: boolean;
  imageURL: string;
  isApplying: boolean;
  isApplied: boolean;
  error: string;
  imageDownloaded: boolean;
  downloadText: string;
  showAccept: boolean;
  siteUrl: string;
  userletters: string;
  sitename: string;
  inclusionpath: string;
  allBadgeImages: string[];
  noBadgesFlag: boolean;
  digitalBadgeScreen: string;
}
export interface IDigitalBadgeProps {
  context: WebPartContext;
  clickcallback: () => void;
  siteUrl: string;
  appTitle: string;
  currentThemeName?: string;
}

export interface IProfileImage {
  url: string;
  width: number;
}

export default class DigitalBadge extends React.Component<IDigitalBadgeProps, IDigitalBadgeState> {
  private columnCount = 0;
  private rowHeight = 0;
  private ROWS_PER_PAGE = 3;
  private MAX_ROW_HEIGHT = 300;
  private commonServiceManager: commonServices;

  constructor(props: IDigitalBadgeProps) {
    super(props);
    let profile: IProfileImage = { url: "", width: 0 };
    //State object Initialization
    this.state = {
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
      downloadText: LocaleStrings.DownloadButtonText,
      userletters: "",
      sitename: siteconfig.sitename,
      inclusionpath: siteconfig.inclusionPath,
      siteUrl: this.props.siteUrl,
      allBadgeImages: [],
      noBadgesFlag: false,
      digitalBadgeScreen: strings.digitalBadgeScreen1
    };
    this.commonServiceManager = new commonServices(this.props.context, this.props.siteUrl);
    this._onDownloadImage = this._onDownloadImage.bind(this);
    this.onUserAcceptance = this.onUserAcceptance.bind(this);
    this.onBadgeSelected = this.onBadgeSelected.bind(this);
    this._onApplyProfileImage = this._onApplyProfileImage.bind(this);
    this.getPhotoBits = this.getPhotoBits.bind(this);
    this.getAllBadgeImages = this.getAllBadgeImages.bind(this);
    this.onRenderCell = this.onRenderCell.bind(this);
    this.getItemCountForPageService = this.getItemCountForPageService.bind(this);
    this.navigateBack = this.navigateBack.bind(this);
  }

  //Component Life cycle method, gets called while the component is getting mounted
  public componentDidMount(): void {
    this._renderListAsync();
  }

  //component life cycle method, gets called whenever the component is updated
  //update aria-label attribute to 'open outlook web application' link in digital badge apply screen
  public componentDidUpdate(prevProps: Readonly<IDigitalBadgeProps>, prevState: Readonly<IDigitalBadgeState>, snapshot?: any): void {
    if (prevState.hasImageSelected !== this.state.hasImageSelected) {
      document.getElementById("linkToChangeProfileImage")?.setAttribute("aria-label", LocaleStrings.digitalBadgeProfileAriaLabel);
    }
  }

  //Get currents user's details from Sharepoint Profiles for Digital Badge processing
  private _renderListAsync() {
    this.props.context.spHttpClient
      .get(
        "/" + this.state.inclusionpath + "/" + this.state.sitename +
        "/_api/SP.UserProfiles.PeopleManager/GetMyProperties",
        SPHttpClient.configurations.v1
      )
      .then((responseuser: SPHttpClientResponse) => {
        responseuser.json().then((datauser: any) => {
          let userassignletters = "";
          let usernamearray = datauser.DisplayName.split(" ");
          if (usernamearray.length === 1) {
            userassignletters = usernamearray[0][0].toUpperCase();
          }
          else if (usernamearray.length > 1) {
            userassignletters =
              usernamearray[0][0].toUpperCase() +
              usernamearray[
                usernamearray.length - 1
              ][0].toUpperCase();
          }
          upn = datauser.Email;
          this.setState({
            showAccept: true,
            userletters: userassignletters,
            isLoading: false,
          });
          this.showUserInformation();
        });
      });
  }

  //Method to get how many items to render per page from specified index
  private getItemCountForPageService = (itemIndex?: number, visibleRect?: IRectangle): number => {
    let serviceObj = this.commonServiceManager.getItemCountForPage(itemIndex, visibleRect, this.MAX_ROW_HEIGHT, this.ROWS_PER_PAGE);
    this.columnCount = serviceObj.columnCount;
    this.rowHeight = serviceObj.rowHeight;
    return serviceObj.itemCountForPage;
  }

  //Multiple Badges section starts
  //Render Fluent UI list cell to show the images and hyperlinks
  private onRenderCell = (item: any, index: number | undefined) => {
    try {
      return (
        <div
          className={dbStyles.listGridTile}
          data-is-focusable
          style={{
            width: 100 / this.columnCount + '%',
          }}
        >
          <div className={dbStyles.listGridSizer}>
            <div className={`${dbStyles.listGridPadder}${!item.enabled ? " " + dbStyles.disabledBadge : ""}`}>
              <a onClick={item.enabled && (() => { this.onBadgeSelected(item.url); })} role="button">
                {this.state.profileImage.url && (
                  <>
                    {this.state.profileImage.url !==
                      "../assets/images/noimage.png" && (
                        <img
                          src={this.state.profileImage.url}
                          className={dbStyles.listGridImage}
                          alt={LocaleStrings.ProfileImageAlt}
                        />
                      )}
                    {this.state.profileImage.url ===
                      "../assets/images/noimage.png" && (
                        <img
                          src={require("../assets/images/noimage.png")}
                          className={dbStyles.listGridImage}
                          alt={LocaleStrings.ProfileImageAlt}
                        />
                      )}
                    <img
                      className={dbStyles.listGridImage}
                      alt={LocaleStrings.BadgeImageAlt}
                      src={item.url}
                      title={item.enabled ? "" : LocaleStrings.BadgePointsTooltip + " " + item.points}
                    />
                    <span className={dbStyles.listGridLabel} title={this.state.userletters} >{this.state.userletters}</span>
                    <span className={dbStyles.listGridLabel} title={item.title} tabIndex={0} role="button"
                      onKeyDown={item.enabled && ((evt: any) => { if (evt.key === strings.stringEnter) this.onBadgeSelected(item.url) })}
                      aria-label={item.title}
                    >
                      {item.title}
                    </span>
                  </>
                )}
              </a>
            </div>
          </div>
        </div>
      );
    }
    catch (error) {
      console.error("CMP_DigitalBadge_onRenderCell \n", JSON.stringify(error));
      this.setState({
        error: strings.TOTErrorMessage + "while displaying the digital badges. Below are the details: \n" + JSON.stringify(error),
        isApplying: false,
        isLoading: false
      });
    }
  }

  public onBadgeSelected(img: string): void {
    try {
      this.setState({
        hasImageSelected: true,
        digitalBadgeScreen: strings.digitalBadgeScreen3,
        imageURL: img,
      });
      this.showUserInformation();
    }
    catch (error) {
      console.error("CMP_DigitalBadge_onBadgeSelected \n", JSON.stringify(error));
      this.setState({
        error: strings.TOTErrorMessage + "while applying the digital badge. Below are the details: \n" + JSON.stringify(error),
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
      //Get member's total points from Event track details list 
      const totalPointsScored = await commonServiceManager.getTotalPointsForMember(upn);

      //Get all badges from the library
      const resultImages: any[] = await commonServiceManager.getAllBadgeImages(strings.DigitalBadgeLibrary, this.props.context.pageContext.user.email.toLowerCase());
      let finalImagesArray: any[] = [];

      if (resultImages.length == 0)
        this.setState({ noBadgesFlag: true, isLoading: false });
      else {
        for (const element of resultImages) {
          if (element.minimumPoints === null ||
            element.minimumPoints === undefined ||
            totalPointsScored >= element.minimumPoints) {
            finalImagesArray.push({
              title: element.title,
              url: element.url,
              enabled: true,
              points: element.minimumPoints
            });
          }
          else {
            finalImagesArray.push({
              title: element.title,
              url: element.url,
              enabled: false,
              points: element.minimumPoints
            });
          }
        }
        //Sorting the badges with enabled first followed by locked badges.
        finalImagesArray.sort((a, b) => {
          if (a.enabled < b.enabled) return 1;
          if (a.enabled > b.enabled) return -1;
        });

        this.setState({ allBadgeImages: finalImagesArray, isLoading: false });
      }
    }
    catch (error) {
      console.error("CMP_DigitalBadge_getAllBadgeImages \n", JSON.stringify(error));
      this.setState({
        error: strings.TOTErrorMessage + "while retrieving the digital badges. Below are the details: \n" + JSON.stringify(error),
        isApplying: false,
        isLoading: false
      });
    }
  }
  //Multiple Badges Section Ends

  //Method to navigate back 
  private navigateBack() {
    try {
      if (this.state.digitalBadgeScreen === strings.digitalBadgeScreen1) {
        this.props.clickcallback();
      }
      else if (this.state.digitalBadgeScreen === strings.digitalBadgeScreen2) {
        this.setState({
          hasAccepted: false,
          showAccept: true,
          digitalBadgeScreen: strings.digitalBadgeScreen1,
          noBadgesFlag: false,
          allBadgeImages: [],
        });
      }
      else if (this.state.digitalBadgeScreen === strings.digitalBadgeScreen3) {
        this.setState({
          hasImageSelected: false,
          isApplied: false
        });
        this.onUserAcceptance();
      }
    }
    catch (error: any) {
      console.error("CMP_DigitalBadge_NavigateBack \n", JSON.stringify(error));
    }
  }

  public render(): React.ReactElement<IDigitalBadgeProps> {
    const isDarkOrContrastTheme = this.props.currentThemeName === strings.themeDarkMode || this.props.currentThemeName === strings.themeContrastMode;
    return (
      <div className={`${dbStyles.digitalBadgeWrapper}${isDarkOrContrastTheme ? " " + dbStyles.digitalBadgeWrapperDarkContrast : ""}`}>
        {this.state.isLoading && (
          <div id="spinnerMessageLabel" aria-label={LocaleStrings.LoadingDigitalBadgeLabel} role="alert"
            aria-live="assertive" className={dbStyles.dbMainSpinner}>
            <Spinner
              size={SpinnerSize.large}
              ariaLabel={LocaleStrings.LoadingSpinnerLabel}
              label={LocaleStrings.LoadingSpinnerLabel}
              aria-hidden="true"
            />
          </div>
        )}
        {!this.state.isLoading && (
          <>
            <div className={dbStyles.dbPanelHeader}>
              <div className={dbStyles.digitalBadgePath}>
                <img src={require("../assets/CMPImages/BackIcon.png")}
                  className={dbStyles.backImg}
                  alt={LocaleStrings.BackButton}
                  aria-hidden="true"
                />
                <span
                  className={dbStyles.backLabel}
                  onClick={this.props.clickcallback}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(evt: any) => { if (evt.key === strings.stringEnter || evt.key === strings.stringSpace) this.props.clickcallback(); }}
                  aria-label={this.props.appTitle}
                >
                  <span title={this.props.appTitle}>
                    {this.props.appTitle}
                  </span>
                </span>
                <span className={dbStyles.border} aria-live="polite" role="alert" aria-label={LocaleStrings.DigitalBadgePageTitle + " Page"} />
                <span className={dbStyles.digitalBadgeLabel}>{LocaleStrings.DigitalBadgePageTitle}</span>
              </div>
            </div>
            <div className={dbStyles.dbPanelBody}>
              <div className={dbStyles.digitalBadge}>
                <div className={`container ${dbStyles.dbContainer}`}>
                  {this.state.isLoading && (
                    <Spinner
                      size={SpinnerSize.large}
                      ariaLabel={LocaleStrings.LoadingSpinnerLabel}
                      label={LocaleStrings.LoadingSpinnerLabel}
                      area-hidden={true}
                    />
                  )}
                  {!this.state.isLoading && (
                    <section aria-live="polite" className={dbStyles.contentSection}>
                      <div className={dbStyles.introPageBox}>
                        {!this.state.hasAccepted && (
                          <div className={dbStyles.divChild1}>
                            <h2>
                              <div className={dbStyles.imgText}>
                                {LocaleStrings.PreAcceptPageTitle}
                              </div>
                            </h2>
                            <img
                              src={require("../assets/CMPImages/AppBanner.png")}
                              className={dbStyles.bannerimage}
                              alt={LocaleStrings.DigitalBadgeAppBannerAltText}
                            />
                          </div>
                        )}
                        <div className={dbStyles.divChild2}>
                          {!this.state.hasAccepted &&
                            this.state.showAccept && (
                              <>
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreAcceptDisclaimer)} />
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreAcceptDisclaimer1)} />
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreAcceptDisclaimer2)} />
                              </>
                            )}
                          <br />
                          {!this.state.hasAccepted &&
                            this.state.showAccept && (
                              <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreAcceptDisclaimer3)} />
                            )}
                          {!this.state.hasAccepted &&
                            !this.state.showAccept && (
                              <div>
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.NotQualifiedPreAcceptDisclaimer)} />
                                <p onClick={this.props.clickcallback}>{LocaleStrings.HowtoGetDigitalBadgeText}</p>
                              </div>
                            )}
                        </div>
                      </div>
                      {this.state.hasAccepted && !this.state.hasImageSelected && (
                        <div className={dbStyles.badgeList}>
                          <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.MultipleBadgeMessage)} />
                          {this.state.noBadgesFlag && (
                            <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.NoBadgeMessage)} />
                          )}
                          <List
                            className={dbStyles.listGrid}
                            items={this.state.allBadgeImages}
                            renderedWindowsAhead={6}
                            getItemCountForPage={this.getItemCountForPageService}
                            getPageHeight={() => this.commonServiceManager.getPageHeight(this.rowHeight, this.ROWS_PER_PAGE)}
                            onRenderCell={this.onRenderCell}
                          />
                        </div>
                      )}
                      {this.state.hasAccepted && this.state.hasImageSelected && (
                        <h2>
                          <div className={dbStyles.badgeDetailsHeading}>{LocaleStrings.DigitalBadgeSubPageTitle}</div>
                        </h2>
                      )}
                      <div className={dbStyles.badgeDetailsContainer}>
                        <div className={dbStyles.badgeBtnArea}>
                          <div className={`${dbStyles.profileContainer} ${dbStyles.profileArea}`}>
                            {this.state.profileImage.url &&
                              this.state.hasAccepted &&
                              this.state.hasImageSelected &&
                              this.state.profileImage.url !==
                              "../assets/images/noimage.png" && (
                                <div style={{ maxWidth: "700px" }} aria-label={LocaleStrings.DigitalBadgeLabel} role="img">
                                  <img
                                    style={{ width: `120px` }}
                                    src={this.state.profileImage.url}
                                    className={dbStyles.profileImage}
                                    alt={LocaleStrings.ProfileImageAlt}
                                    aria-hidden="true"
                                  />
                                  <img
                                    style={{ width: `120px`, marginTop: `-120px` }}
                                    className={dbStyles.badgeImage}
                                    alt={LocaleStrings.BadgeImageAlt}
                                    src={this.state.imageURL}
                                    aria-hidden="true"
                                  />
                                </div>
                              )}
                            {this.state.profileImage.url &&
                              this.state.profileImage.url ===
                              "../assets/images/noimage.png" &&
                              this.state.hasAccepted &&
                              this.state.hasImageSelected && (
                                <div style={{ maxWidth: "700px" }} aria-label={LocaleStrings.DigitalBadgeLabel} role="img">
                                  <img
                                    src={require("../assets/images/noimage.png")}
                                    style={{ width: `120px` }}
                                    className={dbStyles.profileImage}
                                    alt={LocaleStrings.ProfileImageAlt}
                                    aria-hidden="true"
                                  />
                                  <div className={dbStyles.profiletext}>{this.state.userletters}</div>
                                  <img
                                    style={{ width: `120px`, marginTop: `-120px` }}
                                    className={dbStyles.badgeImage}
                                    alt={LocaleStrings.BadgeImageAlt}
                                    src={this.state.imageURL}
                                    aria-hidden="true"
                                  />
                                </div>
                              )}
                            {!this.state.profileImage.url &&
                              this.state.hasAccepted &&
                              this.state.hasImageSelected && (
                                <div>
                                  <img
                                    src={require("../assets/images/noprofile.png")}
                                    id="photoStuff"
                                    alt={LocaleStrings.NoProfileImageAlt}
                                    style={{ width: "150px", height: "auto" }}
                                  />
                                </div>
                              )}
                          </div>
                          {!this.state.isApplying &&
                            this.state.profileImage.url &&
                            this.state.hasAccepted &&
                            this.state.hasImageSelected &&
                            !this.state.isApplied && (
                              <div className={dbStyles.buttonArea}>
                                <PrimaryButton
                                  className={dbStyles.applyBtn}
                                  onClick={this._onApplyProfileImage}
                                  ariaLabel={LocaleStrings.ApplyButtonText}
                                  ariaDescription={LocaleStrings.ApplyButtonAriaDescription}
                                  disabled={
                                    this.state.isApplying ||
                                    this.state.isApplied ||
                                    this.state.error.length > 0
                                  }
                                  title={LocaleStrings.ApplyButton}
                                >
                                  <Icon iconName="Completed" className={dbStyles.acceptIcon} />
                                  {LocaleStrings.ApplyButtonText}
                                </PrimaryButton>
                                <br />
                                {this.state.profileImage.url !==
                                  "../assets/images/noimage.png" && (
                                    <div className={dbStyles.downloadArea}>
                                      <PrimaryButton
                                        iconProps={{ iconName: "Download" }}
                                        className={dbStyles.downloadBtn}
                                        title={this.state.imageDownloaded
                                          ? LocaleStrings.DownloadedButtonSecondaryText
                                          : LocaleStrings.DownloadButtonSecondaryText}
                                        onClick={this._onDownloadImage}
                                        ariaLabel={LocaleStrings.DownloadButtonText}
                                        ariaDescription={LocaleStrings.DownloadButtonAriaDescription}
                                        disabled={
                                          this.state.isApplying ||
                                          this.state.isApplied ||
                                          this.state.imageDownloaded
                                        }
                                      >
                                        {this.state.downloadText}
                                      </PrimaryButton>
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
                              <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreApplyDisclaimer)} />
                            )}
                          {this.state.hasAccepted &&
                            this.state.hasImageSelected &&
                            !this.state.isApplied &&
                            this.state.profileImage.url && (
                              <>
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreApplyDisclaimer1)} />
                                <p dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.PreApplyDisclaimer2)} />
                              </>
                            )}
                        </div>
                      </div>
                      {!this.state.profileImage.url &&
                        this.state.hasAccepted &&
                        this.state.hasImageSelected && (
                          <p className={dbStyles.noProfileDescription} dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.NoProfileImageDescription)} />
                        )}
                      {this.state.isApplying &&
                        !this.state.isApplied && (
                          <div className={dbStyles.applySpinnerContainer}>
                            <Spinner
                              ariaLabel={LocaleStrings.ApplySpinnerLabel}
                              size={SpinnerSize.large}
                              label={LocaleStrings.ApplySpinnerLabel}
                              ariaLive="polite"
                              role="alert"
                            />
                          </div>
                        )}
                      {this.state.isApplied &&
                        !this.state.isApplying && (
                          <div className={dbStyles.messagingContainer}>
                            <MessageBar
                              aria-label={LocaleStrings.DigitalBadgeSuccessMessage}
                              messageBarType={MessageBarType.success}
                              aria-live='polite'
                              role="alert"
                            >
                              <span dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.DigitalBadgeSuccessMessage)} />
                            </MessageBar>
                          </div>
                        )}
                      {this.state.error && (
                        <div className={dbStyles.messagingContainer}>
                          <MessageBar
                            aria-label={this.state.error}
                            messageBarType={MessageBarType.error}
                            aria-live="polite"
                            role="alert"
                          >
                            {this.state.error}
                          </MessageBar>
                        </div>
                      )}
                      <div className={dbStyles.navigateBackAndAcceptBtnWrapper}>
                        <DefaultButton
                          text={LocaleStrings.BackButton}
                          title={LocaleStrings.BackButton}
                          iconProps={{ iconName: 'NavigateBack' }}
                          onClick={this.navigateBack}
                          disabled={this.state.isApplying}
                          className={dbStyles.navigateBackBtn}
                        />
                        {!this.state.isApplying && (
                          <div>
                            {!this.state.hasAccepted &&
                              this.state.showAccept && (
                                <PrimaryButton
                                  className={dbStyles.acceptBtn}
                                  onClick={this.onUserAcceptance}
                                  ariaLabel={LocaleStrings.AcceptButtonText}
                                  ariaDescription={LocaleStrings.AcceptButtonAriaDescription}
                                  title={LocaleStrings.AcceptButtonText}
                                >
                                  <Icon iconName="Completed" className={dbStyles.acceptIcon} />
                                  {LocaleStrings.AcceptButtonText}
                                </PrimaryButton>
                              )}
                            {!this.state.hasAccepted &&
                              !this.state.showAccept && (
                                <p className={dbStyles.unAuthorizedText} dangerouslySetInnerHTML={this.createMarkup(LocaleStrings.UnauthorizedText)} />
                              )}
                          </div>
                        )}
                      </div>
                    </section>
                  )}
                  <canvas
                    id="profileCanvas"
                    width={this.state.profileImage.width}
                    height={this.state.profileImage.width}
                    aria-hidden="true"
                  />
                  <canvas
                    id="profileCanvasDownload"
                    width={this.state.profileImage.width}
                    height={this.state.profileImage.width}
                    aria-hidden="true"
                  />
                </div>
              </div>
            </div>
          </>
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
  public async getPhotoBits(): Promise<any> {
    let canvas: any = document.getElementById("profileCanvas");
    if (canvas.msToBlob) {
      // for IE
      console.log("Function msToBlob found. Using existing function.");
      return Promise.resolve(canvas.msToBlob());
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
        const canvasDownload: any = document.getElementById("profileCanvasDownload");
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
        const canvasDownload: any = document.getElementById("profileCanvasDownload");
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
        //To avoid CORS issue in CDN enabled tenants
        profileImageObj.crossOrigin = "Anonymous";
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
          .then((response) => {
            console.log(response);
            if (response) {
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
      downloadText: LocaleStrings.DownloadingButtonText,
    });
    let canvasDownload: any = document.getElementById("profileCanvasDownload");
    let link: HTMLAnchorElement = document.createElement("a");

    if (canvasDownload.msToBlob) {
      // for IE
      this.setState({ downloadText: LocaleStrings.DownloadedButtonText });
    } else {
      // other browsers
      canvasDownload.toBlob((blob: any) => {
        let url = URL.createObjectURL(blob);
        link.href = url;

        link.setAttribute("download", "myProfileImage.jpg");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        this.setState({ downloadText: LocaleStrings.DownloadedButtonText });
      });
    }
  }

  //On user acceptance show the digital badges for the user to select
  public onUserAcceptance(): void {
    this.setState({
      hasAccepted: true,
      digitalBadgeScreen: strings.digitalBadgeScreen2
    });
    this.getAllBadgeImages();
  }

  public updateUserPhoto(blob: any): Promise<any> {
    let photoPromise: Promise<any> = new Promise(
      (resolve: (arg0: Response) => void, _reject: any) => {
        blob.lastModifiedDate = new Date();
        blob.name = "profile.jpeg";
        this.props.context.msGraphClientFactory
          .getClient('3')
          .then((client: MSGraphClientV3) => {
            client
              .api("me/photo/$value")
              .version("v1.0")
              .header("Content-Type", "image/jpeg")
              .put(blob, (errDb: any, _res: any) => {
                if (!errDb) {
                  resolve(_res);
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
          .getClient('3')
          .then((graphClient: MSGraphClientV3) => {
            graphClient
              .api(graphMyPhotoBitsUrl)
              .version("v1.0")
              .headers({ "Content-Type": "blob", responseType: "blob" })
              .get()
              .then((data: any) => {
                resolve(data);
              })
              .catch((errDb: any) => {
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
          .getClient('3')
          .then((graphClient: MSGraphClientV3) => {
            graphClient
              .api(graphMyPhotoApiUrl)
              .version("v1.0")
              .headers({ "Content-Type": "blob", responseType: "blob" })
              .get()
              .then((data: any) => {
                resolve(data);
              });
          });
      }
    );
    return photoPromise;
  }

  //Get and Process Profile photo data and update Canvas Elements.
  public showUserInformation() {
    let currentProfileImageObj: IProfileImage = { url: "", width: 0 };
    this.getgraphMyPhotoBitsUrl()
      .then((blob) => {
        let blobUrl = URL.createObjectURL(blob);
        currentProfileImageObj.url = blobUrl;
        document.querySelector("#photoStuff")?.setAttribute("src", blobUrl);
        currentProfileImageObj.url = blobUrl;
      })
      .then(() => {
        this.getgraphMyPhotoApiUrl().then((json) => {
          currentProfileImageObj.width = json.width;
          this.setState({ profileImage: currentProfileImageObj });
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
      .substring(1)
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
