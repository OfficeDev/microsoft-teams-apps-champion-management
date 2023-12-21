import React from 'react';
import { WebPartContext } from "@microsoft/sp-webpart-base";

import commonServices from "../Common/CommonServices";
import styles from "../scss/ManageApprovals.module.scss";
import * as LocaleStrings from 'ClbHomeWebPartStrings';
import * as stringsConstants from "../constants/strings";
import siteConfigData from "../config/siteconfig.json";

//Fluent UI Controls
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { Icon } from '@fluentui/react/lib/Icon';
import { Label } from "@fluentui/react/lib/Label";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ITextFieldProps, TextField } from '@fluentui/react/lib/TextField';
import { TOTBreadcrumbLabel } from 'ClbHomeWebPartStrings';

//global variables
let commonServiceManager: commonServices;

export interface IManageConfigSettingsProps {
    context?: WebPartContext;
    siteUrl: string;
    appTitle: string;
    updateAppTitle: Function;
}

export interface IConfigList {
    Id: number;
    ID: number;
    Title: string;
    Value: string;
}
export interface IManageConfigSettingsState {
    showSuccess: boolean;
    showError: boolean;
    errorMessage: string;
    configListSettings: Array<IConfigList>;
    updatedSettings: Array<any>;
    memberListColumns: Array<any>;
    showSpinner: boolean;
    appTitle: string;
    appTitleError: boolean;
}

export default class ManageConfigSettings extends React.Component
    <IManageConfigSettingsProps, IManageConfigSettingsState> {
    constructor(props: IManageConfigSettingsProps) {
        super(props);
        this.state = {
            showSuccess: false,
            showError: false,
            errorMessage: "",
            configListSettings: [],
            updatedSettings: [],
            memberListColumns: [],
            showSpinner: false,
            appTitle: this.props.appTitle,
            appTitleError: false
        };

        //Bind Methods
        this.onToggleSetting = this.onToggleSetting.bind(this);
        this.getConfigListSettings = this.getConfigListSettings.bind(this);
        this.getMemberListColumnNames = this.getMemberListColumnNames.bind(this);
        this.saveConfigSettings = this.saveConfigSettings.bind(this);

        //Create object for CommonServices class
        commonServiceManager = new commonServices(
            this.props.context,
            this.props.siteUrl
        );
    }

    //Getting config settings from config list and column properties from member list
    public async componentDidMount() {
        await this.getConfigListSettings();
        await this.getMemberListColumnNames();
    }

    //Get Config Settings from config list 
    private async getConfigListSettings() {
        try {
            const configListData: IConfigList[] = await commonServiceManager.getAllItemsWithSpecificColumns(
                stringsConstants.ConfigList,
                `${stringsConstants.TitleColumn},${stringsConstants.ValueColumn},${stringsConstants.IDColumn}`);
            if (configListData.length === siteConfigData.configMasterData.length) {
                this.setState({ configListSettings: configListData });
            }
            else {
                this.setState({
                    showError: true,
                    errorMessage:
                        stringsConstants.CMPErrorMessage +
                        ` while loading the page. There could be a problem with the ${stringsConstants.ConfigList} data.`
                });
            }
        }
        catch (error) {
            console.error("CMP_ManageConfigSettings_getConfigListSettings \n", error);
            this.setState({
                showError: true,
                errorMessage:
                    stringsConstants.CMPErrorMessage +
                    `while retrieving the ${stringsConstants.ConfigList} settings. Below are the details: \n` +
                    JSON.stringify(error),
            });
        }
    }

    //Get Member list columns display names
    private async getMemberListColumnNames() {
        try {
            const columnsFilter = "InternalName eq '" + stringsConstants.RegionColumn + "' or InternalName eq '"
                + stringsConstants.CountryColumn + "' or InternalName eq '" + stringsConstants.GroupColumn + "'";
            const columnsDisplayNames: any[] = await commonServiceManager.getColumnsDisplayNames(stringsConstants.MemberList, columnsFilter);
            if (columnsDisplayNames.length > 0) {
                this.setState({ memberListColumns: columnsDisplayNames });
            }
        }
        catch (error) {
            console.error("CMP_ManageConfigSettings_getMemberListColumnNames \n", error);
            this.setState({
                showError: true,
                errorMessage:
                    stringsConstants.CMPErrorMessage +
                    ` while retrieving the ${stringsConstants.MemberList} column data. Below are the details: \n` +
                    JSON.stringify(error),
            });
        }
    }

    //On change of toggle set the states
    private onToggleSetting(_ev: React.MouseEvent<HTMLElement>, checked: boolean, settingName: string) {
        const settings: IConfigList[] = [];
        this.state.configListSettings.forEach((setting: IConfigList) => {
            if (setting.Title === settingName) {
                const updatedSetting: IConfigList = setting;
                updatedSetting.Value = checked ? stringsConstants.EnabledStatus : stringsConstants.DisabledStatus;

                //Add/remove updated settings to/from updatedSettings state
                if (this.state.updatedSettings.find((item: any) => item.id === updatedSetting.ID) !== undefined) {
                    const tempArray = this.state.updatedSettings.filter((item) => {
                        return item.id !== updatedSetting.ID;
                    });
                    this.setState({ updatedSettings: tempArray });
                }
                else {
                    const tempArray = this.state.updatedSettings;
                    const valueFilter = { Value: updatedSetting.Value };
                    const settingObj = { id: updatedSetting.ID, value: valueFilter };
                    tempArray.push(settingObj);
                    this.setState({ updatedSettings: tempArray });
                }
                settings.push(updatedSetting);
            }
            else {
                settings.push(setting);
            }
        });
        this.setState({
            configListSettings: settings, //update config list settings state with updated settings
            showSuccess: false,
            showError: false
        });
    }

    //on change event for App Title field
    private onAppTitleChange(_eve: any, newValue: string) {
        //update the state with new value
        this.setState({
            showSuccess: false,
            showError: false,
            appTitleError: false,
            appTitle: newValue
        })
    }

    //Update the selected settings into the Config list on click of save
    private async saveConfigSettings() {
        let newTitle = this.state.appTitle.trim();
        this.setState({
            showError: false,
            showSuccess: false,
            appTitleError: false,
            errorMessage: "",
            appTitle: newTitle
        })

        if (newTitle === "" || newTitle === undefined) {
            this.setState({
                appTitleError: true
            })
        } else {
            //updating the AppTitle entry in the config list, if it is changed 
            if (this.props.appTitle !== newTitle) {
                let itemID = this.state.configListSettings.filter((item) => {
                    return item.Title === stringsConstants.AppTitle;
                })
                this.state.updatedSettings.push({
                    id: itemID[0].ID,
                    value: { Value: newTitle }
                })
            }

            if (this.state.updatedSettings.length > 0) {
                this.setState({ showSpinner: true });
                commonServiceManager
                    .updateMultipleItemsWithDifferentValues(
                        stringsConstants.ConfigList,
                        this.state.updatedSettings
                    ).then(() => {
                        this.setState({ showSuccess: true, updatedSettings: [], showSpinner: false });
                        //update the state of Parent Component "Manage Approvals" to show the new app title in the app header and breadcrumb
                        if (this.props.appTitle !== newTitle) {
                            this.props.updateAppTitle({
                                appTitle: newTitle
                            });
                        }
                    }).catch((error) => {
                        console.error("CMP_ManageConfigSettings_saveConfigSettings \n", error);
                        this.setState({
                            showError: true,
                            errorMessage:
                                stringsConstants.CMPErrorMessage +
                                " while saving the selection. Below are the details: \n" +
                                JSON.stringify(error),
                            updatedSettings: [],
                            showSpinner: false
                        });
                    });
            }
            else {
                this.setState({
                    showSuccess: true,
                    updatedSettings: []
                });
            }
        }
    }

    //Tooltip for info Icon
    private iconWithTooltip(iconName: string, tooltipContent: string, className: string) {
        return (
            <span className={(styles as any)[className]}>
                <TooltipHost
                    content={tooltipContent}
                    calloutProps={{ gapSpace: 0 }}
                >
                    <Icon iconName={iconName} />
                </TooltipHost>
            </span>
        );
    }

    public render() {
        return (
            <div className={styles.configSettingsContainer}>
                <div>
                    {this.state.showError && (
                        <Label className={styles.errorMessage}>
                            {this.state.errorMessage}
                        </Label>
                    )}
                    {this.state.showSpinner &&
                        <Spinner
                            label={LocaleStrings.ProcessingSpinnerLabel}
                            size={SpinnerSize.large}
                            className={styles.configSettingsSpinner}
                        />
                    }

                    {this.state.configListSettings.length > 0 &&
                        <>
                            {this.state.showSuccess && (
                                <Label className={styles.successMessage} aria-live="polite" role="alert">
                                    <img
                                        src={require('../assets/TOTImages/tickIcon.png')}
                                        alt={LocaleStrings.SuccessIcon}
                                        className={styles.tickImage}
                                    />
                                    {LocaleStrings.ConfigSettingsSaved}
                                </Label>
                            )}
                            <TextField
                                onRenderLabel={() =>
                                    <label className={styles.appNameInputLabel}>
                                        {LocaleStrings.AppTitleLabel}
                                        <span className={styles.asterisk}>*</span>
                                        {this.iconWithTooltip(
                                            "Info", //Icon library name
                                            LocaleStrings.AppTitleInfoIconText,
                                            "configSettingsInfoIcon" //Class name
                                        )}
                                    </label>
                                }
                                maxLength={30}
                                className={styles.appTitleSetting}
                                value={this.state.appTitle}
                                placeholder={LocaleStrings.AppTitlePlaceholderText}
                                onChange={this.onAppTitleChange.bind(this)}
                                errorMessage={this.state.appTitleError && LocaleStrings.AppTitleErrorLabel}
                            />
                            <Toggle
                                checked={this.state.configListSettings.find((setting: IConfigList) => {
                                    return setting.Title === stringsConstants.ChampionEventApprovals;
                                }).Value === stringsConstants.EnabledStatus}
                                label={
                                    <div className={styles.toggleBtnLabel}>
                                        {LocaleStrings.EventApprovalsEnableLabel}
                                        {this.iconWithTooltip(
                                            "Info", //Icon library name
                                            LocaleStrings.EventsApprovalInfoIconTooltipContent,
                                            "configSettingsInfoIcon" //Class name
                                        )}
                                    </div>
                                }
                                inlineLabel
                                defaultChecked
                                onChange={(ev: React.MouseEvent<HTMLElement>, checked: boolean) =>
                                    this.onToggleSetting(ev, checked, stringsConstants.ChampionEventApprovals)}
                                className={styles.configSettingsToggleBtn}
                            />
                            <Label className={styles.listSettingLabel}>
                                {LocaleStrings.ToggleLabelForMemberListColumns}
                                {this.iconWithTooltip(
                                    "Info", //Icon library name
                                    LocaleStrings.TooltipContentForMemberListFieldsHeading,
                                    "configSettingsInfoIcon" //Class name
                                )}
                            </Label>
                            {this.state.memberListColumns.length > 0 &&
                                <>
                                    {this.state.memberListColumns.map((column) => {
                                        const setting: IConfigList = this.state.configListSettings.find((item: IConfigList) => {
                                            return item.Title === column.InternalName;
                                        });
                                        return (
                                            <Toggle
                                                checked={setting.Value === stringsConstants.EnabledStatus}
                                                label={
                                                    <div className={styles.toggleBtnLabel + " " + styles.subLabel} aria-label={`${LocaleStrings.ShowHideToggle} ${column.Title} ${LocaleStrings.ToggleView}`}>
                                                        {column.Title}
                                                    </div>
                                                }
                                                inlineLabel
                                                defaultChecked
                                                onChange={(ev: React.MouseEvent<HTMLElement>, checked: boolean) =>
                                                    this.onToggleSetting(ev, checked, column.InternalName)}
                                                className={styles.configSettingsToggleBtn}
                                            />
                                        );
                                    })}
                                </>
                            }

                            <PrimaryButton
                                text={LocaleStrings.SaveButton}
                                title={LocaleStrings.SaveButton}
                                iconProps={{
                                    iconName: 'Save' //Icon library name
                                }}
                                onClick={this.saveConfigSettings}
                                className={styles.saveBtn}
                            />
                        </>
                    }
                </div>
            </div>
        );
    }
}
