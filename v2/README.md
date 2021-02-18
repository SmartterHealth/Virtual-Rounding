# Virtual Rounding using Microsoft Teams v2
_Updated 4/15/2020_

Version 2 Scripts, Flows, and PowerApps are all ready for use. This documentation will be improved over the next few days.

Version 2 includes the following improved features:
* Ability to invite family/friends
* Unified SharePoint List that contains all rooms and links
   - With views for each sublocation/location still added as tabs in Teams (optional)
* PowerApps for:
   - One click join for Patient Rooms
   >![Teams Policy](/v2/Documentation/Images/PatientJoinApp.png)
   >![Teams Policy](/v2/Documentation/Images/PatientJoinApp2.png)
   - Meeting Join and Configuration for Providers
   >![Teams Policy](/v2/Documentation/Images/VRApp1.png)
   >![Teams Policy](/v2/Documentation/Images/VRApp2.png)
   >![Teams Policy](/v2/Documentation/Images/VRApp3.png)

For deployment assistance, questions or comments, please fill out [this form](https://forms.office.com/Pages/ResponsePage.aspx?id=v4j5cvGGr0GRqy180BHbR6mlTNdIzWRKq7zcu5h9FqNUMVoxSU0yS0hCSVhKMkxRREZaVE1IRU8wVy4u). Someone from Microsoft will reach out as soon as possible.

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
- Hide all apps except the Patient Join PowerApp
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

### Admin Access: 

Access to a Global Administrator account (for Application Consent) 

### Licensing:  

Appropriate licensing for this solution may vary depending on your current agreement and configuration of the virtual rounding solution. Please contact your Microsoft account team for accurate licensing requirement information.  

Depending on the use of agreement, the following is one example of a licensing solution for virtual rounding:  

- Microsoft Teams License for each Patient Room account - minimum of F3 required. 
- Power Automate per flow license for all users who will access the app (not Patient Room accounts). 
- Version 1 of Virtual Rounding is still supported and maintained and does not require this additional licensing. 
- Optional: EM+S licenses for management of Patient Room devices and identities 

# Configuration


## Create Teams Policies

Create Policies in the Microsoft Teams Admin Center matching the below policies. The screenshots below are recommended configuration, but you should configure to your organization's policy/needs.

### Teams Policy
![Teams Policy](/v2/Documentation/Images/TeamsPolicy.png)
### Meeting Policy
![Meeting Policy1](/v2/Documentation/Images/MeetingPolicy1.png)
![Meeting Policy2](/v2/Documentation/Images/MeetingPolicy2.png)
### Live Events Policy
![Live Events Policy](/v2/Documentation/Images/LiveEventsPolicy.png)
### Messaging Policy
![Messaging Policy](/v2/Documentation/Images/MessagingPolicy.png)
### App Permission Policy
![App Permission Policy](/v2/Documentation/Images/AppPermissionPolicy.png)
### App Setup Policy
![App Setup Policy](/v2/Documentation/Images/AppSetupPolicy.png)
### Calling Policy
![Calling Policy](/v2/Documentation/Images/CallingPolicy.png)

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
9. Add the following permissions: Calendars.ReadWrite, Group.ReadWrite.All, OnlineMeetings.ReadWrite.All
10. Click &quot;Grant admin consent for …&quot;
11. From the left menu, click &quot;Certificates &amp; secrets&quot;.
12. Under &quot;Client secrets&quot;, click &quot;+ New client secret&quot;.
13. Provide a description and select an expiry time for the secret and click &quot;Add&quot;.
14. Note down the secret Value.

## SharePoint Site/List Account Setup

In this repository is a PowerShell script (SetupSPO.ps1) that:

1. Creates a SharePoint Site
2. Creates a list
3. Adds custom columns and content type to list.

Before running this script, you will need the following:

- A Global Admin Account
- SharePoint Online PnP PowerShell: [https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps](https://docs.microsoft.com/en-us/powershell/sharepoint/sharepoint-pnp/sharepoint-pnp-cmdlets?view=sharepoint-ps)

### Script

All variables and supporting files will need to be specified in the RunningConfig.json file you can find in this repository. The only varialbe the script needs configured manually is the location of that JSON configuration file.

Once the above is ready, you can run SetupSPO.ps1. As with all open source scripts, please test and review before running in your production environment.

When prompted, sign in with global administrator credentials.


## Patient Room Account Setup

In this repository is a PowerShell script (CreateRooms.ps1) that:

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

## Team/Channel/Tab Creation (Optional)

If you chose to use Teams and Channels to for providers to join/configure meetings/rooms, please see this link.

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

# Power Apps

## Patient Join App
Upload the Patient Join App to PowerApps. The connection to the master SharePoint list in both the PowerApp and the associated flow will need to be updated for proper functionality.

This app then should be published installed on the device.

These instructions will be defined in more detail soon. Please work with a Microsoft Partner in the meantime.

## Virtual Rounding App
Upload the Patient Join App to PowerApps. The connection to the master SharePoint list in the PowerApp and the associated flows will need to be updated for proper functionality.

This app then should be published in Teams and made available to providers.

These instructions will be defined in more detail soon. Please work with a Microsoft Partner in the meantime.

# Security Controls

## Mobile Device Management

We strongly recommend managing the devices with Intune MDM and enabling kiosk mode. More detailed instructions will be added here.

## Conditional Access Policies

We strongly recommend applying a conditional access policy to the Azure AD Group used in _Patient Room Account Setup_ (contains all Patient Room accounts). This policy should limit sign ins to either Intune Managed Devices or specific trusted IPs. This is to limit the risk of the account becoming compromised and a third party logging into an ongoing patient meeting.
