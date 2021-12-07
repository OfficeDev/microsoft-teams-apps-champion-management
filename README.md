# Champion Management Platform 

| [Documentation](https://github.com/OfficeDev/microsoft-teams-apps-champion-management/wiki) | [Solution Overview](https://github.com/OfficeDev/microsoft-teams-apps-champion-management/wiki/Solution-Overview) | [Architecture](https://github.com/OfficeDev/microsoft-teams-apps-champion-management/wiki/Architecture) | [Deployment Guide](https://github.com/OfficeDev/microsoft-teams-apps-champion-management/wiki/Deployment-Guide) | 
| ---- | ---- | ---- | ---- |

## Introduction
The Champion Management Platform is a custom Teams app that enables organizations to onboard and maintain champions/ SME in their organization in Teams, allowing everyone discover Champions right where they collaborate. Utilize this template for multiple scenarios, such as new initiative adoption, champion onboarding, or to maintain organization-wide Subject Matter Experts.

The app provides an easy interface for designated users to add members to the champion program, preview, collaborate and communicate and build a community of power users and coaches.  <br/>

![CMP Screen](./Images/WelcomeCMP.png)                   


## Known issues

1) If site is presenting with spinning blue circle check to ensure permissions are allowed to the ChampionManagementPlatform sharepoint site. If this is happening on the first load experience ensure that the user running first run experience has permissions to create a SharePoint site (first run creates the initial ChampionManagementPlatform site).If you continue to experience the blue circle, please remove the app from Teams and try again. Visit our issues list to log an issue if issue is still persistent. 

<br/>

2) If new users visit the Champion Management Platform app and receive a blue spinning circle, ensure that you have granted contribute/edit permissions to the SharePoint site created during first run to all users (or scoped users accessing the app). The default site created is ChampionManagementPlatform.



## Additional Customization Options

The Champion Management Platform is built to provide a great starting point for managing your program. There are several ways you can modify the solution to fit your needs, with some of the major customizations in this current release coming from modifying list and column options once the solution has been deployed. Common actions include:

Expanding the list of regions / countries / Focus Area / Groups to fit your criteria. Lists provide a very flexible way to provide data points for collection. While we have provided some starter data you will want to modify these values as they are reflected in the Add Members section. This action is done from selecting the list you would like to edit (champion list for this example), selecting the drop down on the column, then navigating to Column settings -> Edit.
<br/>

![Quick Start Guide](./Images/customization.png) 

You could take similar actions in other lists to modify or add in event types and manual counts for activity logging.
<br/>

![Quick Start Guide](./Images/WarningId.png) 
 
## Legal

This app template is provided under the MIT License terms. In addition to these terms, by using this app template you agree to the following:

- You are responsible for complying with all applicable privacy and security regulations related to use, collection, and handling of any personal data by your app. This includes complying with all internal privacy and security policies of your organization if your app is developed to be sideloaded internally within your organization.

- Where applicable, you may be responsible for data related incidents or data subject requests for data collected through your app.

- Any trademarks or registered trademarks of Microsoft in the United States and/or other countries and logos included in this repository are the property of Microsoft, and the license for this project does not grant you rights to use any Microsoft names, logos or trademarks outside of this repository. Microsoft's general trademark guidelines can be found [here](https://www.microsoft.com/en-us/legal/intellectualproperty/trademarks/usage/general.aspx).

- Use of this template does not guarantee acceptance of your app to the Teams app store. To make this app available in the Teams app store, you will have to comply with the [submission and validation process](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish), and all associated requirements such as including your own privacy statement and terms of use for your app.

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA. This project has adopted the Microsoft Open Source Code of Conduct. For more information see the Code of Conduct FAQ or contact opencode@microsoft.com with any additional questions or comments.

## Disclaimer

THIS CODE IS PROVIDED AS IS WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
