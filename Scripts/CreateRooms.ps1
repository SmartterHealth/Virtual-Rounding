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
INSTRUCTIONS:
Please see https://github.com/SmartterHealth/Virtual-Rounding/
#>

#--------------------------Variables---------------------------#
$configFilePath = ".\Scripts\RunningConfig.json"
$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json


$roomListCsvFilePath = $configFile.LocationCsvPaths.Rooms

$groupName = $configFile.TenantInfo.RoomsADGroup

$meetingPolicy = $configFile.TeamsInfo.meetingPolicyName
$messagingPolicy = $configFile.TeamsInfo.messagingPolicyName
$liveEventsPolicy = $configFile.TeamsInfo.liveEventPolicyName
$appPermissionPolicy = $configFile.TeamsInfo.appPermissionPolicyName
$appSetupPolicy = $configFile.TeamsInfo.appSetupPolicyName
$callingPolicy = $configFile.TeamsInfo.callingPolicyName
$teamsPolicy = $configFile.TeamsInfo.teamsPolicyName

#-------------------------Script Setup-------------------------#
#Import-Module AzureAD
Connect-AzureAD
#Connect to Skype for Business Online PowerShell
Import-Module SkypeOnlineConnector
$Session = New-CsOnlineSession
Import-PSSession $Session
#Get ObjectID of Azure AD Group
$groupID = (Get-AzureADGroup -SearchString $groupName).objectID
$accountList = Import-Csv -Path $roomListCsvFilePath

#-----------Create user accounts and apply licensing-----------#
foreach ($account in $accountList){
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $account.AccountPassword
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $upnParts = $account.AccountUPN.Split("@")
    $mailnickname = ($upnParts)[0]
    $upn = $account.AccountUPN.tostring()
    write-host $upn
    #Create Account
    New-AzureADUser -AccountEnabled $true -DisplayName $account.AccountName -UserPrincipalName $upn -Department $account.AccountLocation -UsageLocation "US" -PasswordProfile $PasswordProfile -JobTitle $account.AccountSubLocation -MailNickName $mailnickname
    #Add Account to License Group
    Add-AzureADGroupMember -ObjectId $groupID -RefObjectId (Get-AzureADUser -ObjectId $upn).ObjectID
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