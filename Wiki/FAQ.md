### 1. Can we install in an existing site?

 To maintain permissions and access control, current version of CMP is creating a new site. If you wish to install to a specific existing site you can download the source code and modify the location of SharePoint site the package installs to. This would require a recompile of the package.

### 2. Why is my profile image not updated with Digital Badge?

 This happens when the permissions are not being inherited or approved after deploying package. The users must be able to update their profile images and Graph API permissions must have also been approved during package install. 

 ### 3. I see the below error on Tournament of Teams “Leader Board” 
 
"An unexpected error occured while getting users."

 This happens when the API permission “User.ReadBasic.All” is not approved after upgrading the app from older versions to version 2.0. Refer to the “Upgrade” section and approve the API permission. After approving the permission, it would take some time for it to take effect. 