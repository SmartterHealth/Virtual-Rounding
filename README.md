# Virtual Rounding using Microsoft Teams

_Version: 1.1
Updated 4/8/2020_

For deployment assistance, questions or comments, please fill out [this form](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR6mlTNdIzWRKq7zcu5h9FqNUMVoxSU0yS0hCSVhKMkxRREZaVE1IRU8wVy4u). Someone from Microsoft will reach out as soon as possible.

## Changelog
### Version 1.1
Date: 4/8/2020
* All scripts now have added delays after crucial steps to ensure provisioning of resources, and extra catches to ensure more time is given for provisioning when necessary.
* All scripts no longer need direct modification for variables. A single JSON file is used for all variables, and scripts shouldn't need modifications unless you have desired customizations.
* Bug identified causing meetings to end in the following situation: 
    + Room sitting in meeting -> Provider Joins for a certain period of time -> Provider leaves meeting -> Meeting ends 30 minutes later if no other providers join (only one user in the meeting)
    + A new part of the Virtual Rounding solution has been added to solve this bug. There is now a free meeting bot you can deploy to always be joined to the meeting and serve as a constant second meeting participant to ensure the 30 minute timer does not apply.
    + Please note that there are new API permissions required in the Azure AD App registration to support this solution.

## Overview

This is the Virtual Rounding solution referenced in the Microsoft Health &amp; Life Sciences [blog post](https://aka.ms/teamsvirtualrounding). Please see that blog post for more an overview of the use case. This repository serves as the technical documentation.

## Disclaimer

_This solution is a sample and may be used with Microsoft Teams for dissemination of reference information only. This solution is not intended or made available for use as a medical device, clinical support, diagnostic tool, or other technology intended to be used in the diagnosis, cure, mitigation, treatment, or prevention of disease or other conditions, and no license or right is granted by Microsoft to use this solution for such purposes. This solution is not designed or intended to be a substitute for professional medical advice, diagnosis, treatment, or judgement and should not be used as such. Customer bears the sole risk and responsibility for any use. Microsoft does not warrant that the solution or any materials provided in connection therewith will be sufficient for any medical purposes or meet the health or medical requirements of any person._

## Solution Design

A device will be deployed in each patient room needed. These will be referred to as &quot;patient rooms&quot; in this documentation. That device will be locked down in Kiosk mode to the Microsoft Teams application.

Each patient room will have an associated Office 365 account, with only Microsoft Teams and Exchange Online licensing applied. Custom Teams Policies will be applied to the accounts to limit the capabilities within the Teams application &amp; meetings, including:

- Disable Chat
- Disable Calling
- Disable Organization browsing
- Disable Meeting &amp; Live Event creation
- Disable Discovery of Private Teams
- Disable Installation/Adding Apps
- Hide all apps except Calendar
- Disable Meeting Features: Meet Now, Cloud Recording, Transcription, Screen Sharing, PowerPoint Sharing, Whiteboard, Shared Notes, Anonymous user admission, Meeting Chat

Each patient room will have an ongoing Teams meeting running for a long period of time (months or longer), and that meeting will be reused for that room as patients flow in and out of rooms. As noted in the known limitations, there is a 24 hour timeout; Please see that section for guidance.

Doctors will not be directly invited to any meetings, but instead have access to a Team or Teams with a list of meetings pinned as a tab (from SharePoint). Doctors will be able to join a Patient Room meeting via the Join URL hyperlink in the list.

## Known Limitations &amp; Warnings

- Patient Room accounts will be able to browse and join public Teams. Limit the presence of those in your directory or deploy Information Barriers to prevent this.
- Patient Room accounts can technically create Teams if this is not already restricted. Consider implementing restrictions to Team creation to these accounts to limit this ability if that is a concern.
- While Patients Room accounts are prevented from exposing PHI during meetings (no chat, whiteboard, or shared notes access), Doctors do not have those same limitations (unless you choose to apply custom meeting policies to Doctors as well). Ensure Doctors have proper training or documentation to _not_ use those features of put PHI in them. Any content posted in those features will be visible to the next patient in the room.
- If a patient goes to the show participants list, they are technically able to invite other users from your directory. There is no current workaround for this besides training and patient supervision.
- If a patient taps/clicks on the doctor's name while in the meeting, they can see Azure AD profile information for that Doctor. There is no current workaround for this besides training and patient supervision. Some hospitals have used generic workstations with generic Teams logins to get around this for the doctors.
- This solution is built with cloud only Azure AD Accounts in mind for the Patient Room accounts. Any variation from that will have to be coded manually.
- This solution relies on familiarity with PowerShell, Azure AD Admin, Teams Admin, and Power Automate (formerly known as Microsoft Flow). and may require customization for your specific environment. Please fill out [this form](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR6mlTNdIzWRKq7zcu5h9FqNUMVoxSU0yS0hCSVhKMkxRREZaVE1IRU8wVy4u) to contact us if you need assistance.
- A single participant cannot be joined to a meeting for more than 24 hours. The devices will need to rejoin the meetings at least once every 24 hours. To avoid service interruptions, consider training floor staff that is already in the room once daily to end and rejoin the meeting.

## Prerequisites

- Access to a Global Administrator account (for Application Consent)
- A service account with a Power Automate Premium license
  - If not available, PowerShell can be used instead of Power Automate (this will require customization not available in this repository)
- Enough Office 365 Licenses for each Patient Room account (any License SKU that includes Microsoft Teams and Exchange Online [plan 1 or 2]; If needed, please contact your Microsoft Account team for the E1 trial available during COVID-19)
- Optional: EM+S licenses for management of Patient Room devices and identities

# Configuration

All configuration steps below assume that you would like to set this up at scale with a large amount of accounts. If you would like to test out the solution or POC with a smaller amount of user accounts, no scripting or Power Automate is needed. Simply follow along but skip running scripts/flows and instead manually complete the same steps that are listed for each script/flow.

## Create Teams Policies

Create Policies in the Microsoft Teams Admin Center matching the below policies. The screenshots below are recommended configuration, but you should configure to your organization's policy/needs.

### Teams Policy
![Teams Policy](/Documentation/Images/TeamsPolicy.png)
### Meeting Policy
![Meeting Policy1](/Documentation/Images/MeetingPolicy1.png)
![Meeting Policy2](/Documentation/Images/MeetingPolicy2.png)
### Live Events Policy
![Live Events Policy](/Documentation/Images/LiveEventsPolicy.png)
### Messaging Policy
![Messaging Policy](/Documentation/Images/MessagingPolicy.png)
### App Permission Policy
![App Permission Policy](/Documentation/Images/AppPermissionPolicy.png)
### App Setup Policy
![App Setup Policy](/Documentation/Images/AppSetupPolicy.png)
### Calling Policy
![Calling Policy](/Documentation/Images/CallingPolicy.png)

## Application Registration

For various steps in this process we will need to call the Microsoft Graph. To do that, an app registration is required in Azure AD. This will require a Global Administrator account.

1. Navigate to [https://aad.portal.azure.com/#blade/Microsoft\_AAD\_IAM/ActiveDirectoryMenuBlade/RegisteredApps](https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps) and sign in as a Global Administrator.
2. Click New Registration.
3. Provide an application name, select &quot;Accounts in this organizational directory only&quot;, and leave Redirect URI blank. Click Register.
4. Note down Application and Directory IDs to use later.
5. From the left menu, click &quot;API permissions&quot; to grant some permissions to the application.
6. Click &quot;+ Add a permission&quot;.
7. Select &quot;Microsoft Graph&quot;.
8. Select Application permissions.
9. Add the following permissions: Calendars.ReadWrite, Calls.InitiateGroupCall.All, Calls.JoinGroupCall.All, Group.ReadWrite.All, OnlineMeetings.ReadWrite.All
10. Click &quot;Grant admin consent for …&quot;
11. From the left menu, click &quot;Certificates &amp; secrets&quot;.
12. Under &quot;Client secrets&quot;, click &quot;+ New client secret&quot;.
13. Provide a description and select an expiry time for the secret and click &quot;Add&quot;.
14. Note down the secret Value.

## Patient Room Account Setup

In this repository is a PowerShell script that:

1. Creates the accounts
2. Adds the account to a group (for tracking and group based licensing)
3. Applies Custom Teams Policies

Before running this script, you will need the following:

- An Azure AD Security Group
  - This group should be empty, and will only be used for the patient room accounts. If provisioning manually, ensure all users are added to this group. It will be used later for licensing and in the flows.
  - You will also need to setup Group Based Licensing for the Azure AD Security Group. Please assign an Office 365 license to the group and disable all assignment options except for Microsoft Teams, Skype for Business and Exchange Online. Detailed instructions for Group Based Licensing can be found here: [https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-groups-assign](https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-groups-assign).
- A CSV file with the desired Patient Room account information
  - Columns:
    - AccountName
      - Desired Room Name. Ensure this is easily identifiable to your clinical staff.
    - AccountUPN
      - Desired UserPrincipalName (email address) of the room.
      - This will only be used for login to the Teams application on the device.
      - Do not use a domain that is federated (ADFS, Ping, etc), as this will be a cloud only account.
      - Must not conflict with any other accounts in your directory.
    - AccountPassword
      - Must comply with your Azure AD Password Policies
    - AccountLocation
      - Building or Location of the room. This will be used to categorize the rooms later on and will be tied to a Team name (a suffix can be added) (see _Team/List/Tab_ sections below).
      - This will be filled into the "Department" field of the user.
    - AccountSubLocation
      - Sub Location of the room (example: Floor 1). This will be used to categorize the rooms later on and must be tied to either a Channel name (see _Team/List/Tab_ sections below).
      - This will be filled into the "Job Title" field of the user.
  - Sample file available (RoomAccounts.csv)
- Azure AD PowerShell Module v2: [https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
- Skype for Business Online PowerShell: [https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell](https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell)

### Script

All variables and supporting files will need to be specified in the RunningConfig.json file you can find in this repository. The only varialbe the script needs configured manually is the location of that JSON configuration file.

Once the above is ready, you can run CreateRooms.ps1. As with all open source scripts, please test and review before running in your production environment.

When prompted, sign in with administrator credentials that are able to create Azure AD accounts and assign Teams policies.

## Team/List/Tab Creation

Depending on your setup, you may want one Team or multiple Teams for Doctors to use to navigate and join the Patient Room meetings.

We recommend a single Team with Channels for each location involved so that Doctors are able to join any room at any location during a Time of crisis. This is the method that will be covered and supported in this guide.

If doctor&#39;s should only be able to join rooms at specific locations (hospitals/clinics), we recommend separate teams per location, or a single Team with _Private_ Channels for each location involved. This guide does not cover that method at this time however, and you will need to adapt to your needs. This will be added to the guide at a different time.

In SharePoint, we will be leveraging SharePoint lists to store and surface the meeting join links. This guide will cover creating multiple SharePoint lists, one for each location, and having them added as Tabs to the associated channel. Each list will also get a custom view applied.

### Script

In this repository is a PowerShell Script (CreateTeamsAndSPO.ps1) that:

1. Creates the Team
2. Sets Team Settings:
  i. Visibility: Private
  ii. Disables member capabilities: Add/Remove Apps, Create/Update/Remove Channels, Create/Update/Remove Connectors, Create/Update/Remove Tabs
3. Adds members to Team
4. Creates SharePoint Lists in the associated SharePoint site
5. Adds columns and custom view to lists
   i. Columns: Title (exists by default), RoomLocation, RoomSubLocation, MeetingLink, EventID(skip if provisioning manually).
   ii. View: (Create a new view)[https://support.office.com/en-us/article/Create-change-or-delete-a-view-of-a-list-or-library-27AE65B8-BC5B-4949-B29B-4EE87144A9C9] and then (add in the JSON)[https://support.microsoft.com/en-us/office/formatting-list-views-f737fb8b-afb7-45b9-b9b5-4505d4427dd1?ui=en-us&rs=en-us&ad=us] from SharePointViewFormatting.json
6. Creates Channels and pins the SharePoint list as a Tab
7. Removes Wiki Tabs

Before running this script, you will need the following:

- Azure AD Security Groups
  - You will specify security groups to copy membership from to the individual Teams in the below CSV (_MembersGroupName_).
  - These groups should contain the provider's accounts that you want to be added to the Teams as members. Do not include any of the room accounts you created earlier. They should _not_ be members of the Team. 
- The App ID and Client Secret from the Azure AD App Registration (earlier step in this guide).
- A CSV file with the desired Team(s) information
  - Columns:
    - LocationName
      - Location Name. This must match the location names used for AccountLocation in _Patient Room Account Setup_. Ensure all Location Names from that earlier script are represented.
      - You will be able to add a suffix to this to make a more readable Team name by using a variable in the script
    - MembersGroupName
      - Name of an Azure AD Group (or synced AD Group) containing the members to be added to the Team.
      - These groups should contain the provider's accounts that you want to be added to the Teams as members. Do not include any of the room accounts you created earlier. They should _not_ be members of the Team. 
  - Sample file available (LocationList.csv)
- A second CSV file with the desired Channel(s)/List(s)
  - Columns:
    - LocationSubName
      - Sub Location Name. This must match the location names used for AccountSubLocation in _Patient Room Account Setup_. Ensure all Sub Location Names from that earlier script are represented.
    - LocationName
      - Location Name. This must match the location names used for AccountLocation in _Patient Room Account Setup_. Ensure all Location Names from that earlier script are represented.
  - Sample file available (SubLocationList.csv)
- Azure AD PowerShell Module v2: [https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
- Microsoft Teams PowerShell: [https://www.powershellgallery.com/packages/MicrosoftTeams/](https://www.powershellgallery.com/packages/MicrosoftTeams/)
- SharePoint Online PnP PowerShell: [https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)

All variables and supporting files will need to be specified in the RunningConfig.json file you can find in this repository. The only varialbe the script needs configured manually is the location of that JSON configuration file.

Once the above is ready, you can run CreateTeamsAndSPO.ps1. As with all open source scripts, please test and review before running in your production environment.

## Patient Room Meeting Setup

## Meeting Creation

To create the meetings, we will use Power Automate. Power Automate offers a simple way to call the Microsoft Graph API, and the ability to run on a regular basis if we need in the future.

Prerequisites:

- A Power Automate Premium license will be required for this piece (P1, P2, Per User or Per App all work).
- An account with the Power Automate license applied to it, used for creating the Flows (ideally a service account).
- SetupMeetingsFlow.zip from this repository
- Get the Group GUID/ObjectID for your Azure AD Group used in _Patient Room Account Setup_ (find in the group properties in the Azure AD Portal)

Instructions:

1. Login to flow.microsoft.com
2. Click on &quot;My flows&quot;.
3. Click &quot;Import&quot;.
4. Upload SetupMeetingsFlow.zip
5. Update all variables, the SharePoint Site base URL in the final step of the flow, and the Group ID.

Once it's been at least 3 hours since you've created the room accounts, you can run the Flow to create all the meeting links. Ideally, wait at least 24 hours. This is to ensure the Teams Policies properly apply to the room accounts before a meeting is created.

## Meeting Bot
This section covering the meeting bot is _draft_, and we recommend reaching out to your Microsoft Partner or account team for assistance with this. We will finalize this section over the next 48 hours as we continue to build.

A meeting bot can be used to get around the 30 minute timeout issue mentioned in the changelog at the top of this page. The meeting bot will sit in each meeting and serve as a second meeting participant to avoid the 30 minute timeout (which starts as soon as a meeting is down to one participant). The bot is subject to the same 30 minute and 24 hour timeouts that standard accounts have. Therefore, it is crucial that patient device not hang up the meeting, as that would leave the bot as the lone participant in the meeting, starting the 30 minute timer.
The meeting bot is joined into a meeting using a Graph API call, which can be automated using Power Automate or PowerShell to ensure it rejoins every 24 hours, and potentially sooner depending on your needs. The below will outline the basics of the bot setup process. Ensure you have updated your Azure AD App Registration with the newly added API permissions before starting.

### Bot Configuration
1. Go to https://dev.botframework.com/bots/new
2. Fill out all the pertinent information, ensuring to use the app ID from your Azure AD App registration.
3. Add Microsoft Teams as a channel
4. Select the calling tab, and select the checkbox to _Enable calling_. For your webhook, enter any https URL. We will never be calling this bot, so this field won't be relevant, but it is required to enter something.
5. In Microsoft Teams, select Apps from the left pane and then select App Studio.
6. From the top pane, click Manifest editor and then Create a new app from the left pane.
7. In the App details tab, provide the basic information.
8. Navigate to the Capabilities section, and select the Bots tab. Then select Set Up in the right pane.
9. Fill in the desired bot name
10. Select the Select from one of my existing bots option, and find your bot from above in the dropdown.
11. Check all options under Calling Bot and Scope and press Save
12. Use app studio to deploy the bot to your tenant.

### Adding the bot to a Teams meeting
A Graph API call using your Azure AD App Registration (Client ID, Client Secret, Tenant ID) will allow us to add the bot to an existing scheduled meeting.
To get the items that the API call will need, get your meeting join link, which should look like this:

`https://teams.microsoft.com/l/meetup-join/19%3ameeting_YWNiYzA2NTctOGIzMy00MzRhLTkyNmUtZGY4NzM2YTFhNmEz%40thread.v2/0?context=%7b%22Tid%22%3a%226be58f7f-c45d-43f9-89e4-b97ec2a06d8e%22%2c%22Oid%22%3a%22ac2ea2ab-9845-4308-a99c-8fdc6548ceac%22%7d`

Decoding that URI, we get this:

`https://teams.microsoft.com/l/meetup-join/19:meeting_YWNiYzA2NTctOGIzMy00MzRhLTkyNmUtZGY4NzM2YTFhNmEz@thread.v2/0?context={"Tid":"6be58f7f-c45d-43f9-89e4-b97ec2a06d8e","Oid":"ac2ea2ab-9845-4308-a99c-8fdc6548ceac"}`

The two items we need from the decoded uri are:

- threadId: `19:meeting_YWNiYzA2NTctOGIzMy00MzRhLTkyNmUtZGY4NzM2YTFhNmEz@thread.v2`
- organizerId `ac2ea2ab-9845-4308-a99c-8fdc6548ceac`

Using that information, call the graph API using the below to add the bot to the meeting:

Call: `POST https://graph.microsoft.com/beta/communications/calls`

Body:
```json
{
  "@odata.type": "#microsoft.graph.call",
  "callbackUri": "INSERT URI FROM STEP 4 ABOVE",
  "tenantId": "INSERT TENANTID HERE",
  "meetingInfo": {
    "@odata.type": "#microsoft.graph.organizerMeetingInfo",
    "organizer": {
      "@odata.type": "#microsoft.graph.identitySet",
      "user": {
        "@odata.type": "#microsoft.graph.identity",
        "id": "INSERT ORGANIZERID HERE",
        "tenantId": "INSERT TENANTID HERE"
      }
    },
    "allowConversationWithoutHost": true
   },
  "mediaConfig": {
    "@odata.type": "#microsoft.graph.serviceHostedMediaConfig"
    },
   "chatInfo": {
    "@odata.type": "#microsoft.graph.chatInfo",
    "threadId": "INSERT THREADID HERE",
    "messageId": "0"
  }
}
```
### Flow for automating adding bot to meetings
We have built a flow for you to use to continuously add the bot to all meetings. This will ensure the bot is not outside of the meeting for more than 25 minutes, ensuring the patient room will not be kicked out either. Again, the only reason the bot will be kicked out is after 24 hours, or if the patient room is not joined 30 minutes. The flow will solve both of those.

Prerequisites:

- A Power Automate Premium license will be required for this piece (P1, P2, Per User or Per App all work).
- An account with the Power Automate license applied to it, used for creating the Flows (ideally a service account).
- SetupMeetingsFlow.zip from this repository
- Get the Group GUID/ObjectID for your Azure AD Group used in _Patient Room Account Setup_ (find in the group properties in the Azure AD Portal)

Instructions:

1. Login to flow.microsoft.com
2. Click on &quot;My flows&quot;.
3. Click &quot;Import&quot;.
4. Upload AddBotToMeetingsFlow.zip
5. Update all variables.

## Meeting Updating

If there is an error with a meeting link, there is a flow that can be manually run to update the link. Please note when this happens, someone will need to end the meeting on the patient room device and join the new meeting.
Please check back here soon for the details of the flow for this purpose.

# Security Controls

## Mobile Device Management

We strongly recommend managing the devices with Intune MDM and enabling kiosk mode. More detailed instructions will be added here.

## Conditional Access Policies

We strongly recommend applying a conditional access policy to the Azure AD Group used in _Patient Room Account Setup_ (contains all Patient Room accounts). This policy should limit sign ins to either Intune Managed Devices or specific trusted IPs. This is to limit the risk of the account becoming compromised and a third party logging into an ongoing patient meeting.