### Upgrade to version 2.5 from 2.4, 2.3, 2.2, 2.1, 2.0, 1.3 and 1.2

If you are already having version 2.4, 2.3, 2.2, 2.1, 2.0, 1.3 or 1.2 installed on your tenant follow the below steps to upgrade to version 2.5 with a global admin account:

 NOTE: If you are using/seeing modern app catalog refer to the go to the [Modern App Catalog](#modern-app-catalog) section.

1.  Navigate to App Catalog with a tenant Admin account. Do not delete the existing package. Upload the new package "cmp.sppkg" that is downloaded from "sharepoint/solution" folder which will replace the existing package.  

    ![Upgrade](../Images/Upgrade-1.png) 

1. Click on "Deploy".

    ![Upgrade](../Images/Upgrade-2.png) 

1. "Check In" the package.

    ![Upgrade](../Images/Upgrade2.0-3.png) 

1. Select the package and click "Sync to Teams" from the ribbon and wait for the "Successfully synced to teams solution" message to appear.

    ![SyncToTeams](../Images/SyncToTeams.png) 

1. Repeat steps 1-3 to deploy a new package file "mgt-spfx-vv3.sppkg" to the App catalog. 
The package can be downloaded from "sharepoint/solution" folder.This package is required for the person card feature when hover over a champion name in the app.

    ![Quick Start Guide](../Images/GraphtoolkitModern.png) 

1. Navigate to the ***SharePoint admin center***. Under expand the ***Advanced*** menu in the left navigation and select ***API access***. Select and approve the below pending requests associated with ***championmanagement***.

    **User.Read.All** : Allows the app to read the full set of profile properties, reports, and managers of other users in your organization, on behalf of the signed-in user. CMP uses this permission to display person cards for all champions.

    **People.Read**: Allows the app to read a scored list of people relevant to the signed-in user. The list can include local contacts, contacts from social networking or your organization's directory, and people from recent communications. CMP uses this permission to display person cards for all champions.

    ![Quick Start Guide](../Images/APIAccess-Upgrade.png)  

1. The Champion Management Platform will be updated to the latest version and you will see changes reflected in Teams. Please note that if you do not see changes reflected in Teams after 30 minutes you can log out and back in and clear the Teams cache to see changes immediately. 

1. This steps applies only if you are upgrading from versions 2.1 or older. This step is not required if you are upgrading from version 2.2. <br>
If you already had "Tournament of Teams" enabled before the upgrade, click on "Enable Tournament of Teams" under "Admin Tools" section again. That will provision new lists for Tournament Reporting feature. "Tournament of Teams" icon will not be visible in the "Get Started" section without this step.

    ![Upgrade](../Images/Upgrade-3.png) 

    **NOTE:** If there is any completed tournament before the upgrade, enabling Tournament of Teams will take couple of minutes (depends on number of completed tournaments) to load the data about completed tournaments into the Tournaments Report and Participants Report lists. Do not navigate away from the screen until it is completed with success message.

    ![Upgrade](../Images/EnableTOT.png) 

    ![Upgrade](../Images/TOTScreen.png) 

1. The API permission "Sites.Manage.All" can be removed from "API Permissions" in SharePoint admin portal only if the permission is not used in any other apps in your tenant. The upgraded package for CMP is not using this permission anymore.

    ![Upgrade](../Images/Upgrade-4.png) 

1. This steps applies only if you are upgrading from version 2.1. A new CMP logo has been added in this package. If you have not customized the app logo for the CMP application, you can delete the "CMP Logo" library from SharePoint site so that it will be automatically re-created when the app is launched and the MS logo will be replaced with new CMP logo image in the library.

### Modern App Catalog 

``` This section applies only if you are using/seeing modern app catalog.```

1. Navigate to App Catalog. Do not delete the existing package. Upload the new package that is downloaded from "sharepoint/solution" folder which will replace the existing package.

    ![Upgrade](../Images/Upgrade_Modern_AppCatalog1.png) 

1. After uploading the package, select "Enable this app and add it to all sites" and click on "Enable App"

    ![Upgrade](../Images/Modern_AppCatalog2.png)

1. Skip this step.

    ![Upgrade](../Images/Modern_AppCatalog3.png)

1. Once done, click on "Add to Teams" to make this app available in Teams. Some times you may see an error after clicking "Add to Teams". You can still wait for few minutes and check the app to confirm if you see the new features.

    ![Upgrade](../Images/Modern_AppCatalog4.png)

1. Repeat steps 1-2 to deploy a new package file "mgt-spfx-vv3.sppkg" to the App catalog by selecting "Enable this app and it to all sites" option. This package can be downloaded from "sharepoint/solution" folder. This package is required for the person card feature when hover over a champion name in the app.

    ![Quick Start Guide](../Images/GraphtoolkitModern.png)     

``` Continue with steps 6 to 10 from the previous section.```

### Upgrade to version 2.5 from 1.1

If you are already having 1.1 installed on your tenant and want to upgrade to 2.5 the existing app and SharePoint site 'ChampionManagementSite' have to be deleted. 

If you have current members and events you will want to export those list items and import re-import them into the respective list areas. We have expanded our lists to have some additional data as well so you may need to populate additional fields. 

- Memberlist (same information) 
- EventList (same information) 
- EventTrackDetails (added two colums to contain event name + member name) 

Follow the below steps to upgrade your install and optionally also export and import the data you may already have for the program. 

1.	Delete the existing App from Teams.

    ![Quick Start Guide](../Images/Upgrade1.png) 

2.	Delete 'Members List' from root site. <br/>
    a. Remember to export members if you are wanting to import the members back in

3.	Delete 'ChampionManagementPlatform' Sharepoint site from both 'Active Sites' and 'Deleted Sites'.

    a. Remember to export the Event List and Event Track List if you are wanting to import the events details back in. 

    ![Quick Start Guide](../Images/Upgrade2.png) 
 
4. Delete 'cmp.sppkg' from App Catalog.	
5. Wait for around 30 minutes and startover the installation of new package following instructions from the section 'Deployment-Guide'.
6. Once the install has completed and first run is done, you can visit the site assets of the ChampionManagementPlatform and import any of the exported Champion data