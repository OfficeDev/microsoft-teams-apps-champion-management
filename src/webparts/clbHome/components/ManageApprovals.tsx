import { IPivotItemProps, Pivot, PivotItem } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import React, { Component } from 'react';
import commonServices from '../Common/CommonServices';
import styles from "../scss/ManageApprovals.module.scss";
import ApproveChampion from './ApproveChampion';
import ChampionsActivities from './ChampionsActivities';
import ManageConfigSettings from './ManageConfigSettings';
import * as stringsConstants from "../constants/strings";

//declaring common services object
let commonServiceManager: commonServices;
export interface IManageApprovalsProps {
    context: WebPartContext;
    siteUrl: string;
    onClickBack: Function;
    isPendingChampionApproval: boolean;
    isPendingEventApproval: boolean;
    appTitle: string;
    updateAppTitle: Function;
    currentThemeName?: string;
}
export interface IManageApprovalsState {
    isPendingChampionApproval: boolean;
    isPendingEventApproval: boolean;
    appTitle: string;
}
export default class ManageApprovals extends Component<IManageApprovalsProps, IManageApprovalsState> {

    constructor(props: IManageApprovalsProps) {
        super(props);
        this.state = {
            isPendingChampionApproval: this.props.isPendingChampionApproval,
            isPendingEventApproval: this.props.isPendingEventApproval,
            appTitle: this.props.appTitle
        };

        this.setState = this.setState.bind(this);

        //Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );
    }


    //Updating the state of parent component whenever this component is updated 
    public componentDidUpdate(prevProps: Readonly<IManageApprovalsProps>, prevState: Readonly<IManageApprovalsState>, snapshot?: any): void {
        //updating state of the parent component 'ClbHome" to show the new app title in app header and breadcrumb
        if (prevState.appTitle !== this.state.appTitle) {
            this.props.updateAppTitle({
                appTitle: this.state.appTitle
            });
        }
    }

    _customRenderer(
        link?: IPivotItemProps,
        defaultRenderer?: (link?: IPivotItemProps) => JSX.Element | null,
    ): JSX.Element | null {
        if (!link || !defaultRenderer) {
            return null;
        }
        return (
            <span>
                <span>&nbsp;&nbsp;{link.headerText}</span>
                <img src={require(`../assets/CMPImages/BellIcon.svg`)} className={styles.indicatorIcon} />
                {defaultRenderer({ ...link, headerText: undefined })}
            </span>
        );
    }

    _customRendererNoIcon(
        link?: IPivotItemProps,
        defaultRenderer?: (link?: IPivotItemProps) => JSX.Element | null,
    ): JSX.Element | null {
        if (!link || !defaultRenderer) {
            return null;
        }
        return (
            <span>
                <span>&nbsp;&nbsp;{link.headerText}&nbsp;&nbsp;</span>
                {defaultRenderer({ ...link, headerText: undefined })}
            </span>
        );
    }


    public render() {
        const isDarkOrContrastTheme = this.props.currentThemeName === stringsConstants.themeDarkMode || this.props.currentThemeName === stringsConstants.themeContrastMode;
        return (
            <div className={`container ${styles.manageApprovalsContainer}${isDarkOrContrastTheme ? " " + styles.manageApprovalsContainerDarkContrast : ""}`}>
                <div className={styles.manageApprovalsPath}>
                    <img src={require("../assets/CMPImages/BackIcon.png")}
                        className={styles.backImg}
                        alt={LocaleStrings.BackButton}
                        aria-hidden="true"
                    />
                    <span
                        className={styles.backLabel}
                        onClick={() => { this.props.onClickBack(); }}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(evt: any) => { if (evt.key === stringsConstants.stringEnter) this.props.onClickBack(); }}
                        aria-label={this.props.appTitle}
                    >
                        <span title={this.props.appTitle}>
                            {this.props.appTitle}
                        </span>
                    </span>
                    <span className={styles.border}></span>
                    <span className={styles.manageApprovalsLabel}>{LocaleStrings.AdminTasksLabel}</span>
                </div>
                <Pivot
                    linkFormat='tabs'
                    className={styles.manageApprovalsPivot}
                    defaultSelectedKey="0"
                >
                    <PivotItem
                        headerText={LocaleStrings.ChampionsListPageTitle}
                        itemKey="0"
                        onRenderItemLink={this.state.isPendingChampionApproval ? this._customRenderer : this._customRendererNoIcon}
                        ariaLabel={navigator.userAgent.match(/Android/i) ? `${LocaleStrings.ChampionsListPageTitle} 1 of 3` : ""}
                    >
                        <span title={LocaleStrings.ChampionsListPageTitle}>
                            <ApproveChampion
                                context={this.props.context}
                                siteUrl={this.props.siteUrl}
                                setState={this.setState}
                            />
                        </span>
                    </PivotItem>
                    <PivotItem
                        headerText={LocaleStrings.ChampionActivitiesLabel}
                        itemKey="1"
                        onRenderItemLink={this.state.isPendingEventApproval ? this._customRenderer : this._customRendererNoIcon}
                        ariaLabel={navigator.userAgent.match(/Android/i) ? `${LocaleStrings.ChampionActivitiesLabel} 2 of 3` : ""}
                    >
                        <span title={LocaleStrings.ChampionActivitiesLabel}>
                            <ChampionsActivities
                                context={this.props.context}
                                siteUrl={this.props.siteUrl}
                                setState={this.setState}
                            />
                        </span>
                    </PivotItem>
                    <PivotItem
                        headerText={LocaleStrings.ManageConfigSettingsLabel}
                        itemKey="2"
                        ariaLabel={navigator.userAgent.match(/Android/i) ? `${LocaleStrings.ManageConfigSettingsLabel} 3 of 3` : ""}
                    >
                        <span title={LocaleStrings.ManageConfigSettingsLabel}>
                            <ManageConfigSettings
                                context={this.props.context}
                                siteUrl={this.props.siteUrl}
                                appTitle={this.props.appTitle}
                                updateAppTitle={this.setState}
                            />
                        </span>
                    </PivotItem>
                </Pivot>
            </div>
        );
    }
}
