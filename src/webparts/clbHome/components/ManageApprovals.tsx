import { IPivotItemProps, Pivot, PivotItem, PivotLinkFormat } from '@fluentui/react';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import React, { Component } from 'react';
import commonServices from '../Common/CommonServices';
import styles from "../scss/ManageApprovals.module.scss";
import ApproveChampion from './ApproveChampion';
import ChampionsActivities from './ChampionsActivities';
import ManageConfigSettings from './ManageConfigSettings';

//declaring common services object
let commonServiceManager: commonServices;
export interface IManageApprovalsProps {
    context: WebPartContext;
    siteUrl: string;
    onClickBack: Function;
    isPendingChampionApproval: boolean;
    isPendingEventApproval: boolean;
}
export interface IManageApprovalsState {
    isPendingChampionApproval: boolean;
    isPendingEventApproval: boolean;
}
export default class ManageApprovals extends Component<IManageApprovalsProps, IManageApprovalsState> {

    constructor(props: IManageApprovalsProps) {
        super(props);
        this.state = {
            isPendingChampionApproval: this.props.isPendingChampionApproval,
            isPendingEventApproval: this.props.isPendingEventApproval,
        };

        this.setState = this.setState.bind(this);

        //Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );
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
        return (
            <div className={`container ${styles.manageApprovalsContainer}`}>
                <div className={styles.manageApprovalsPath}>
                    <img src={require("../assets/CMPImages/BackIcon.png")}
                        className={styles.backImg}
                        alt={LocaleStrings.BackButton}
                    />
                    <span
                        className={styles.backLabel}
                        onClick={() => { this.props.onClickBack(); }}
                        title={LocaleStrings.CMPBreadcrumbLabel}
                    >
                        {LocaleStrings.CMPBreadcrumbLabel}
                    </span>
                    <span className={styles.border}></span>
                    <span className={styles.manageApprovalsLabel}>{LocaleStrings.AdminTasksLabel}</span>
                </div>
                <Pivot
                    linkFormat={PivotLinkFormat.tabs}
                    className={styles.manageApprovalsPivot}
                    defaultSelectedKey="0"
                >
                    <PivotItem
                        headerText={LocaleStrings.ChampionsListPageTitle}
                        itemKey="0"
                        headerButtonProps={{ title: LocaleStrings.ChampionsListPageTitle }}
                        onRenderItemLink={this.state.isPendingChampionApproval ? this._customRenderer : this._customRendererNoIcon}
                    >
                        <ApproveChampion
                            context={this.props.context}
                            siteUrl={this.props.siteUrl}
                            setState={this.setState}
                        />
                    </PivotItem>
                    <PivotItem
                        headerText={LocaleStrings.ChampionActivitiesLabel}
                        itemKey="1"
                        headerButtonProps={{ title: LocaleStrings.ChampionActivitiesLabel }}
                        onRenderItemLink={this.state.isPendingEventApproval ? this._customRenderer : this._customRendererNoIcon}
                    >
                        <ChampionsActivities
                            context={this.props.context}
                            siteUrl={this.props.siteUrl}
                            setState={this.setState}
                        />
                    </PivotItem>
                    <PivotItem
                        headerText={LocaleStrings.ManageConfigSettingsLabel}
                        itemKey="2"
                        headerButtonProps={{ title: LocaleStrings.ManageConfigSettingsLabel }}
                    >
                        <ManageConfigSettings
                            context={this.props.context}
                            siteUrl={this.props.siteUrl}
                        />
                    </PivotItem>
                </Pivot>
            </div>
        );
    }
}
