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
Please see https://aka.ms/virtualroundingcode
#>

#-------------------Configurable Variables---------------------#
$configFilePath = ".\GitHub\Virtual-Rounding\v2\Scripts\RunningConfig.json"

#--------------System Variables (DO NOT MODIFY)----------------#
$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json

$roomListCsvFilePath = $configFile.LocationCsvPaths.Rooms

$roomsGroupName = $configFile.TenantInfo.RoomsADGroup

$meetingPolicy = $configFile.TeamsInfo.meetingPolicyName
$messagingPolicy = $configFile.TeamsInfo.messagingPolicyName
$liveEventsPolicy = $configFile.TeamsInfo.liveEventPolicyName
$appPermissionPolicy = $configFile.TeamsInfo.appPermissionPolicyName
$appSetupPolicy = $configFile.TeamsInfo.appSetupPolicyName
$callingPolicy = $configFile.TeamsInfo.callingPolicyName
$teamsPolicy = $configFile.TeamsInfo.teamsPolicyName

$useMFA = $configFile.TenantInfo.MFARequired
$adminUPN = $configFile.TenantInfo.GlobalAdminUPN

#-------------------------Script Setup-------------------------#
Function Test-Existence {
    [CmdletBinding()]
    param(
        $value,
        $errorMsg
    )
    try{
        if($value){
            return $true
        }
        else{
            throw $errorMsg
        }
    }
    catch{
        Write-Error  $_
    }
}
#-------------------------Script Setup-------------------------#
if (!$useMFA) {$creds = Get-Credential -Message 'Please sign in to your Global Admin account:' -UserName $adminUPN}

Test-Existence((Get-Module AzureAD),'The AzureAD Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module AzureAD
if ($useMFA) {Connect-AzureAD -ErrorAction Stop}
else {Connect-AzureAD -Credential $creds -ErrorAction Stop}

Test-Existence((Get-Module SkypeOnlineConnector),'The SkypeOnlineConnector Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module SkypeOnlineConnector

$roomsGroupID = (Get-AzureADGroup -Filter "DisplayName eq '$roomsGroupName'").objectID
$existingGroupMembers = Get-AzureADGroupMember -ObjectId $roomsGroupID
$accountList = Import-Csv -Path $roomListCsvFilePath -ErrorAction Stop

#-----------Create user accounts and apply licensing-----------#
foreach ($account in $accountList){
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = $account.AccountPassword
    $PasswordProfile.ForceChangePasswordNextLogin = $false
    $upnParts = $account.AccountUPN.Split("@")
    $mailnickname = ($upnParts)[0]
    $upn = $account.AccountUPN.tostring()
    $userCheck = $null
    $userCheck = Get-AzureADUser -Filter "UserPrincipalName eq '$upn'"
    #Test for Account
    if ($null -eq $userCheck) {
        #Create Account
        New-AzureADUser -AccountEnabled $true -DisplayName $account.AccountName -UserPrincipalName $upn -Department $account.AccountLocation -UsageLocation "US" -PasswordProfile $PasswordProfile -JobTitle $account.AccountSubLocation -MailNickName $mailnickname
        #Add Account to License Group
        $userObjectID = (Get-AzureADUser -ObjectId $upn).ObjectID
        try {Add-AzureADGroupMember -ObjectId $roomsGroupID -RefObjectId $userObjectID}
        catch {
            if ($existingGroupMembers -contains $userObjectID) {Write-Host "$upn is already a member of $roomsGroupName group. Continuing..." -ForegroundColor DarkYellow}
            else {Write-Host "Unable to add $upn to $roomsGroupName group. The script will need to be restarted." -ErrorAction Stop -ForegroundColor Red}
        }
        Write-Host "Created $upn and added to $roomsGroupName group." -ForegroundColor Green
    }
    else {
        throw "$upn already exists in tenant. Skipped adding to Azure AD Group. WARNING: Teams policies will be applied to this account. Cancel run if this is unexpected."
    }
}
#Wait for licensing application and Teams/Exchange provisioning
Write-Host "Script will now pause for 15 minutes to allow for licensing application and Teams/Exchange provisioning of new accounts. Check https://admin.microsoft.com/AdminPortal/Home#/teamsprovisioning to verify status." -ForegroundColor Green
Start-Sleep -Seconds 900 #15 minutes

#---------------------Apply Teams Policies---------------------#
#Connect to Skype for Business Online PowerShell
if ($useMFA) {$skypeSession = New-CsOnlineSession -UserName $adminUPN -ErrorAction Stop}
else {$skypeSession = New-CsOnlineSession -Credential $creds -ErrorAction Stop}
Import-PSSession $skypeSession -ErrorAction Stop

foreach ($account in $accountList){
    $upn = $account.AccountUPN
    #Check if account is ready
    $user = Get-CsOnlineUser -Identity $upn -ErrorAction SilentlyContinue
    while ($null -eq $user){
        Write-Host "$upn is not ready for Teams Policies. Would you like to wait 15 more minutes (w), skip this user (s), or cancel the script (c)? (Default is Wait)" -ForegroundColor Yellow
        $readHost = Read-Host " ( w / s / c )"
        Switch ($readHost){
            W {Write-host "Wait 15 more minutes"; Start-Sleep -Seconds 900} 
            S {Write-Host "Skip $upn"; $skip = $true; Continue} 
            C {Write-Host "Cancel Script"; break} 
            Default {Write-Host "Waiting 15 more minutes" -ForegroundColor Green; Start-Sleep -Seconds 900}
        }
        if($skip){Continue}
        $user = Get-CsOnlineUser -Identity $upn -ErrorAction SilentlyContinue
    }
    if($skip){Continue}
    Grant-CsTeamsAppPermissionPolicy -Identity $upn -PolicyName $appPermissionPolicy
    Grant-CsTeamsAppSetupPolicy -Identity $upn -PolicyName $appSetupPolicy
    Grant-CsTeamsCallingPolicy -Identity $upn -PolicyName $callingPolicy
    Grant-CsTeamsMeetingBroadcastPolicy -Identity $upn -PolicyName $liveEventsPolicy
    Grant-CsTeamsMeetingPolicy -Identity $upn -PolicyName $meetingPolicy
    Grant-CsTeamsMessagingPolicy -Identity $upn -PolicyName $messagingPolicy
    Grant-CsTeamsChannelsPolicy -Identity $upn -PolicyName $teamsPolicy
    Grant-CsTeamsUpgradePolicy -Identity $upn -PolicyName UpgradeToTeams #Sets account to Teams Only mode
}

Write-Host "Script Complete." -ForegroundColor Green