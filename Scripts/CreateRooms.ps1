<#
DISCLAIMER: 
----------------------------------------------------------------
This sample is provided as is and is not meant for use on a production environment.
It is provided only for illustrative purposes. The end user must test and modify the
sample to suit their target environment. 

Microsoft can make no representation concerning the content of this sample. Microsoft
is providing this information only as a convenience to you. This is to inform you that
Microsoft has not tested the sample and therefore cannot make any representations 
regarding the quality, safety, or suitability of any code or information found here.    
----------------------------------------------------------------
#>

<#
PREREQUISITE POWERSHELL MODULES:
-Azure AD Powershell v2: https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0
-Skype for Business Online Powershell: https://docs.microsoft.com/en-us/office365/enterprise/powershell/manage-skype-for-business-online-with-office-365-powershell

INSTRUCTIONS:
Prepare a CSV file of desired accounts to be created for patient rooms/devices (Columns expected: AccountName, AccountUPN, AccountPassword, AccountLocation).
Create an Azure AD Security Group and apply proper group based licensing.
Create custom Policies in the Teams Admin Center.
Define variables below, and run the script.
When prompted, sign in with Administrator credentials (with the ability to create Azure AD accounts & assign Teams policies).
#>


#--------------------------Variables---------------------------#
#Path of the CSV file
#Columns expected: AccountName, AccountUPN, AccountPassword, AccountLocation, AccountSubLocation
$csvFile = ""
#Name of Security Group for Group Based Licensing
$groupName = ""
#Name of Teams Policies configured to be applied to accounts
$meetingPolicy = ""
$messagingPolicy = ""
$liveEventsPolicy = ""
$appPermissionPolicy = ""
$appSetupPolicy = ""
$callingPolicy = ""
$teamsPolicy = ""

#-------------------------Script Setup-------------------------#
#Import-Module AzureAD
Connect-AzureAD
#Connect to Skype for Business Online PowerShell
Import-Module SkypeOnlineConnector
$Session = New-CsOnlineSession
Import-PSSession $Session
#Get ObjectID of Azure AD Group
$groupID = (Get-AzureADGroup -SearchString $groupName).objectID
$accountList = Import-Csv -Path $csvFile

#-----------Create user accounts and apply licensing-----------#
foreach ($account in $accountList){
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $account.Password
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $upnParts = $account.AccountUPN.Split("@")
    $mailnickname = ($upnParts)[0]
    #Create Account
    New-AzureADUser -AccountEnabled $true -DisplayName $account.AccountName -UserPrincipalName $account.AccountUPN -Department $account.AccountLocation -UsageLocation "US" -PasswordProfile $PasswordProfile -JobTitle $account.AccountSubLocation -MailNickName $mailnickname
    #Add Account to License Group
    Add-AzureADGroupMember -ObjectId $groupID -RefObjectId (Get-AzureADUser -ObjectId $account.UserPrincipalName)
}
#Wait for licensing application and Teams/Exchange provisioning
Start-Sleep -Seconds 900 #15 minutes

#---------------------Apply Teams Policies---------------------#
foreach ($account in $accountList){
    Grant-CsTeamsAppPermissionPolicy -Identity $account.AccountUPN -PolicyName $appPermissionPolicy
    Grant-CsTeamsAppSetupPolicy -Identity $account.AccountUPN -PolicyName $appSetupPolicy
    Grant-CsTeamsCallingPolicy -Identity $account.AccountUPN -PolicyName $callingPolicy
    Grant-CsTeamsMeetingBroadcastPolicy -Identity $account.AccountUPN -PolicyName $liveEventsPolicy
    Grant-CsTeamsMeetingPolicy -Identity $account.AccountUPN -PolicyName $meetingPolicy
    Grant-CsTeamsMessagingPolicy -Identity $account.AccountUPN -PolicyName $messagingPolicy
    Grant-CsTeamsChannelsPolicy -Identity $account.AccountUPN -PolicyName $teamsPolicy
    Grant-CsTeamsUpgradePolicy -Identity $account.AccountUPN -PolicyName UpgradeToTeams #Account should be in Teams Only mode
}