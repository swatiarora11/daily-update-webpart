# Daily Update Webpart - Overview and Deployment Guide

**SPFx Updates Webpart** is a client side web part built using Sharepoint Framework (SPFx). This webpart has integration with a sharepoint list to pull prefed updates and display them on given dates. This webpart can also be used by admins to send activity feed in teams to keep the employees engaged and informed about the latest updates.

The webpart leverages the tight integration between SharePoint Online and Microsoft Teams to support scenarios like daily updates, announcements, relevant news, thought of the day, corporate communication etc.

## Prerequisites 

To begin, you will need:

* An active subscription of Sharepoint Online and Microsoft Teams.
* Office 365 account(s) with administrative rights for Sharepoint Online and Microsoft Teams workloads.

## Anatomy of Webpart

An overview snapshot of the **SPFx Updates Webpart** highlighting its major sections is shown in the below figure. 
<img src="screenshots/webpart overview.png"/>
Let us dive into few details to understand configuration of this webpart and how it controls the content being shown in various sections of the webpart. Each webpart section and corresponding fields of property pane configuration which affect webpart presentation and functionality is explained below.

1. **Picture** - sourced from the url provided in the **Picture Url** field of webpart property pane
2. **Title** - sourced from the text entered in the **Title** field of webpart property pane
3. **Sub Title** - sourced from the shrepoint list provided using **Select sites** and **Select a list** fields of webpart property pane
4. **Description** - sourced from the shrepoint list provided using **Select sites** and **Select a list** fields of webpart property pane
5. **Send Feed Button** - 
    * enabled only for the user whose UPN is filled in **Admin UPN** field of webpart property pane
    * On click of the **Send Feed Button**, webpart sends teams activity feed to members of the O365 group selected in the **Select Group** field of webpart property pane
    * activity feed will be successfully sent only if Teams application with App ID filled in the **Teams App ID** field of webpart property pane is installed in user's personal scope.

All the mappings of webpart sections to corresponding configuration fields are summarized in the snapshot below.
<img src="screenshots/webpart anatomy.jpeg"/>

## Deployment Guide

Though **SPFx Updates Webpart** can be used in any Sharepoint Online site, this deployment guide explains the steps required for deployment of this webpart with the help of an example Viva Connections application which is deployed in Sharepoint Online and also available in personal scope of various users of Microosoft Teams.

### Step 1. Setup Viva Connections

Refer the documentation available [here](https://docs.microsoft.com/en-us/viva/connections/viva-connections-overview) to get step-by-step guidance on how to setup Viva Connections.

### Step 2. Create SharePoint List
* Open **Site Contents** page on the Viva Connections site. Select **New -> List** then select **Blank List** and enter the name of the List as "Chairman Speak".
* Rename the **Title** column to **Author** and add **Live**, **Update** columns to the list and select the type as shown below. <img src="screenshots/sharepoint list-settings.png"/>
**Note:** Do not create columns with name **ID**, **Title**, **Created by** and **Modified by** as they exist by default in the list <img src="screenshots/Sharepoint Site-list.png"/>

### Step 3. Upload SPPKG to Sharepoint App Catalog
1. Download the [Sharepoint Solution Package](https://github.com/swatiarora11/spfx-daily-feed/blob/main/sppkg/daily-update-webpart.sppkg) file from this repository and save the file to your computer.

2. For uploading, go to **Sharepoint Admin Center -> More features -> Apps -> App Catalog -> Apps for Sharepoint**.
Upload this file into the **App Catalog** by selecting **Upload**, browsing the file in the downloaded folder and then selecting **Deploy**. <img src="screenshots/App catalog-sharepoint.png"/>
You will see that **SPFx Update Webpart** is now listed in the **App Catalog**.

### Step 4. Grant API Permissions
Once the app package is uploaded, navigate to **API Access** page in **Sharepoint Admin Center** and approve the below permissions.
<img src="screenshots/API Access in SP.png"/>

### Step 5. Create and Install Teams App
1.	Navigate to **App Registrations** in **Azure Portal** and copy the **Application(client) ID** of **Sharepoint Online Client Extensibility Web application Principal** app. We will need this later for installation of Teams App. 
<img src="screenshots/Azureportal, webapplicationid.png"/>

2.	Download [teams.zip](https://github.com/swatiarora11/spfx-daily-feed/blob/main/sppkg/daily-update-webpart.sppkg) file from this git repository and extract the same to a local folder.
3.	Change following fields in "developer" section of the downloaded **manifest.json** file to values as appropriate for your organization.
* name
* websiteUrl
* privacyUrl
* termsOfUseUrl
4.	Change the "id" field under "webApplicationInfo" section in the manifest to "Application (client) ID" value of **Sharepoint Online Client Extensibility Web application Principal** as copied in serial 1 above and save **manifest.json** file. 
5.	Create a ZIP package with **manifest.json** and app icon files (**color.png** and **outline.png**). Make sure that there are no nested folders within this ZIP package.
6.	Navigate to Microsoft Teams Admin Center. Under **Teams apps > Manage apps** section, click **+ Upload** and upload ZIP package file created in the previous step. Once upload is complete, you will be able to see the Connections app under the **Manage apps** tab as shown below.<img src="screenshots/manage teams-teams admin center.png"/>
7.	Ensure that Custom App policy permission has been enabled under **Permission Policies**
8.	Now add this app to **App Setup Policies**, which in turn will make the app visible to all users in Microsoft Teams canvas. To add this for all users, select **Global Policy**.
<img src="screenshots/App setup policy-teams admin center.png"/>

9.	Now set the sequence to make the app visible to each user. We recommend to pin the app in the top 5, so that it is easily visible to end users on each client. Hit **Save** to make this change.

### Step 6. Publishing the Webpart
1. Edit the Sharepoint page where you would like to add the webpart.
2. Click **Add a new section (+)** on the left-hand side of the page, and then click **One Column**.
3. Click **+**, then select the **Update** web part.
<img src="screenshots/Add Spfx webpart .png"/>

4. Click on **Edit Webpart** and configure the webpart as explained in **Anatomy of Webpart** section above and Republish.
<img src="screenshots/Edit the webpart-description.png"/>