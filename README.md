# Virtual Rounding using Microsoft Teams

_Version: 0.2
Updated 3/23/2020_
_Please check back here end of day 3/23/2020. Documentation will be updated and supporting files/scripts for large sclae provisioning will be available at that time. In the meantime, manual provisioning is recommended._

## Overview

This is the Virtual Rounding solution referenced in the Microsoft Health &amp; Life Sciences [blog post](https://aka.ms/teamsvirtualrounding). Please see that blog post for more details on the use case. This repository serves as the technical documentation.

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

Each patient room will have an ongoing Teams meeting running all day, and that meeting will be reused for that room as patients flow in and out of rooms. Every 24 hours, the meeting link will be updated (using Power Automate).

Doctors will not be directly invited to any meetings, but instead have access to a Team or Teams with a list of meetings pinned as a tab (from SharePoint). Doctors will be able to join a Patient Room meeting via the Join URL hyperlink in the list.

## Known Limitations &amp; Warnings

- Patient Room accounts will be able to browse and join public Teams. Limit the presence of those in your directory or deploy Information Barriers to prevent this.
- Patient Room accounts can technically create Teams if this is not already restricted. Consider implementing restrictions to Team creation to these accounts to limit this ability if that is a concern.
- While Patients Room accounts are prevented from exposing PHI during meetings (no chat, whiteboard, or shared notes access), Doctors do not have those same limitations (unless you choose to apply custom meeting policies to Doctors as well). Ensure Doctors have proper training or documentation to _not_ use those features of put PHI in them. Any content posted in those features will be visible to the next patient in the room.
- If a patient goes to the show participants list, they are technically able to invite other users from your directory. There is no current workaround for this.
## Prerequisites

- Access to a Global Administrator account (for Application Consent)
- A service account with a Power Automate Premium license
  - If not available, PowerShell can be used instead of Power Automate
- Enough Office 365 Licenses for each Patient Room account (any License SKU that includes Microsoft Teams and Exchange Online [plan 1 or 2])
- Optional: Intune licenses for management of Patient Room devices

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

## Application Registration {screenshots to be added}

For various steps in this process we will need to call the Microsoft Graph. To do that, an app registration is required in Azure AD. This will require a Global Administrator account.

1. Navigate to [https://aad.portal.azure.com/#blade/Microsoft\_AAD\_IAM/ActiveDirectoryMenuBlade/RegisteredApps](https://aad.portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps) and sign in as a Global Administrator.
2. Click New Registration.
3. Provide an application name, select &quot;Accounts in this organizational directory only&quot;, and leave Redirect URI blank. Click Register.
4. Note down Application and Directory IDs to use later.
5. From the left menu, click &quot;API permissions&quot; to grant some permissions to the application.
6. Click &quot;+ Add a permission&quot;.
7. Select &quot;Microsoft Graph&quot;.
8. Select Application permissions.
9. Add the following permissions:
  1. Calendars.ReadWrite
  2. OnlineMeetings.ReadWrite.All
  3. Group.ReadWrite.All
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
  - This group should be empty, and will only be used for the patient room accounts
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
      - Building or Location of the room. This will be used to categorize the rooms later on and must be tied to either a Team name or Channel name (see _Team Creation_ and _Channel Setup_ sections below).
  - Sample file located [here]
- Azure AD PowerShell Module v2: [https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
- Skype for Business Online PowerShell: [https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell](https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell)

You will also need to setup Group Based Licensing for the Azure AD Security Group. Please assign an Office 365 license to the group and disable all assignment options except for Microsoft Teams and Exchange Online. Detailed instructions for Group Based Licensing can be found here: [https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-groups-assign](https://docs.microsoft.com/en-us/azure/active-directory/users-groups-roles/licensing-groups-assign).

### Script

Once the above is ready, you can run CreateRooms.ps1. As with all open source scripts, please test and review before running in your production environment. Ensure you fill in the appropriate variables before running the script.

There will be two sign in prompts during the script. Sign in with administrator credentials that are able to create Azure AD accounts and assign Teams policies.

## Team/List/Tab Creation

Depending on your setup, you may want one Team or multiple Teams for Doctors to use to navigate and join the Patient Room meetings.

We recommend a single Team with Channels for each location involved so that Doctors are able to join any room at any location during a Time of crisis. This is the method that will be covered and supported in this guide.

If doctor&#39;s should only be able to join rooms at specific locations (hospitals/clinics), we recommend separate teams per location, or a single Team with _Private_ Channels for each location involved. This guide does not cover that method at this time however, and you will need to adapt to your needs. This will be added to the guide at a different time.

In SharePoint, we will be leveraging SharePoint lists to store and surface the meeting join links. This guide will cover creating multiple SharePoint lists, one for each location, and having them added as Tabs to the associated channel. Each list will also get a custom view applied.

### Script

In this repository is a PowerShell Script (CreateTeams.ps1) that:

1. Creates the Team
2. Sets Team Settings:
  1. Visibility: Private
  2. Disables member capabilities: Add/Remove Apps, Create/Update/Remove Channels, Create/Update/Remove Connectors, Create/Update/Remove Tabs
3. Adds members to Team
4. Creates SharePoint Lists in the associated SharePoint site
5. Adds columns and custom view to lists
6. Creates Channels and pins the SharePoint list as a Tab
7. Removes Wiki Tabs

Before running this script, you will need the following:

- An Azure AD Security Group
  - This group should be empty, and will only be used for the patient room accounts
- The App ID and Client Secret from the Azure AD App Registration (earlier step in this guide)
- A CSV file with the desired Team(s) information
  - Columns:
    - LocationName
      - Location Name. This must match the location names used for AccountLocation in _Patient Room Account Setup_. Ensure all Location Names from that earlier script are represented.
      - You will be able to add a suffix to this to make a more readable Team name by using a variable in the script
    - MembersGroupName
      - Name of an Azure AD Group (or synced AD Group) containing the members to be added to the Team.
  - Sample file located [here]
- A second CSV file with the desired Channel(s)/List(s)
  - Columns:
    - SubLocationName
      - Sub Location Name. This must match the location names used for AccountSubLocation in _Patient Room Account Setup_. Ensure all Sub Location Names from that earlier script are represented.
    - LocationName
      - Location Name. This must match the location names used for AccountLocation in _Patient Room Account Setup_. Ensure all Location Names from that earlier script are represented.
  - Sample file located [here]
- Azure AD PowerShell Module v2: [https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0](https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
- Microsoft Teams PowerShell: [https://www.powershellgallery.com/packages/MicrosoftTeams/](https://www.powershellgallery.com/packages/MicrosoftTeams/)
- SharePoint Online PnP PowerShell: [https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)

Once the above is ready, you can run CreateTeamsAndSPO.ps1. As with all open source scripts, please test and review before running in your production environment. Ensure you fill in the appropriate variables before running the script.

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
5. Update all variables (more details to be added here soon)

## Meeting Updating

Since meeting links can only last for 24 hours, we will use Power Automate to create new ones and update the SharePoint lists every 24 hours.
Please check back here soon for the details of this flow.

# Security Controls

## Mobile Device Management

We strongly recommend managing the devices with Intune MDM and enabling kiosk mode. More detailed instructions will be added here.

## Conditional Access Policies

We strongly recommend applying a conditional access policy to the Azure AD Group used in _Patient Room Account Setup_ (contains all Patient Room accounts). This policy should limit sign ins to either Intune Managed Devices or specific trusted IPs. This is to limit the risk of the account becoming compromised and a third party logging into an ongoing patient meeting.

(instructions to be added)
