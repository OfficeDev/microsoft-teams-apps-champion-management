### Upgrade to version 2.0 from 1.3 and 1.2

If you are already having version 1.3 or 1.2 installed on your tenant follow the below steps to upgrade to version 2.0:

1.  Navigate to App Catalog with a tenant Admin account. Do not delete the existing  1.3 or 1.2 package. Upload the new package that is downloaded from "sharepoint/solution" folder which will replace the existing package.  

![Upgrade 1.3](../Images/Upgrade2.0-1.png) 

2. Click on "Deploy".

![Upgrade 1.3](../Images/Upgrade2.0-2.png) 

3. "Check In" the package.

![Upgrade 1.3](../Images/Upgrade2.0-3.png) 

4. Select the package and click "Sync to Teams" from the ribbon and wait for the "Successfully synced to teams solution" message to appear

![Upgrade 1.3](../Images/SyncToTeams.png) 

5. Navigate to SharePoint admin center. Under Advanced menu in the left navigation select API access. Select and approve the additional pending request associated with championmanagement as shown below. 
After approving the permission, it would take some time for it to take effect. 

![Upgrade 1.3](../Images/Upgrade2.0-4.png) 

6. The Champion Management Platform will be updated to the latest version and you will see changes reflected in Teams. Please note that if you do not see changes reflected in Teams after 30 minutes you can log out and back in and clear the Teams cache to see changes immediately. 

### Upgrade to version 2.0 from 1.1

If you are already having 1.1 installed on your tenant and want to upgrade to 2.0 the existing app and SharePoint site 'ChampionManagementSite' have to be deleted. 

If you have current members and events you will want to export those list items and import re-import them into the respective list areas. We have expanded our lists to have some additional data as well so you may need to populate additional fields. 

Memberlist (same information) 
EventList (same information) 
EventTrackDetails (added two colums to contain event name + member name) 

Follow the below steps to upgrade your install and optionally also export and import the data you may already have for the program. 

1.	Delete the existing App from Teams.
<br/>

![Quick Start Guide](../Images/Upgrade1.png) 

 
2.	Delete 'Members List' from root site. <br/>
    a. Remember to export members if you are wanting to import the members back in 
3.	Delete 'ChampionManagementPlatform' Sharepoint site from both 'Active Sites' and 'Deleted Sites'.

    a. Remember to export the Event List and Event Track List if you are wanting to import the events details back in. 

<br/>

![Quick Start Guide](../Images/Upgrade2.png) 
 
4. Delete 'cmp.sppkg' from App Catalog.	
5. Wait for around 30 minutes and startover the installation of new package following instructions from the section 'Deployment-Guide'.
6. Once the install has completed and first run is done, you can visit the site assets of the ChampionManagementPlatform and import any of the exported Champion data