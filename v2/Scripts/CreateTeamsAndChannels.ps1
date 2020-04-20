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
#>

<#
INSTRUCTIONS:
Please see https://aka.ms/virtualroundingcode
#>

#--------------------------Variables---------------------------#
$configFilePath = "C:\Users\mafritz\OneDrive - Microsoft\Documents\GitHub\Virtual-Rounding\v2\Scripts\RunningConfig.json"
$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json

$locationsCsvFile = $configFile.LocationCsvPaths.Locations
$subLocationsCsvFile = $configFile.LocationCsvPaths.SubLocations
$groupOwner = $configFile.TenantInfo.MeetingSchedulingUser
$spviewJsonFilePath = $configFile.ViewJson.SPViewJsonFilePath
$teamNameSuffix = $configFile.GroupConfiguration.RoundingTeamSuffix
$clientId = $configFile.ClientCredential.Id
$clientSecret = $configFile.ClientCredential.Secret
$tenantName = $configFile.TenantInfo.TenantName
$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl
$sharepointMasterSiteName = $configFile.TenantInfo.SPOMasterSiteName
$sharepointMasterListName = $configFile.TenantInfo.SPOMasterListName
$tabName = $configFile.TeamsInfo.TabName
$shareTabName = $configFile.TeamsInfo.ShareTabName

$useMFA = $configFile.TenantInfo.MFARequired
$adminUPN = $configFile.TenantInfo.GlobalAdminUPN

#--------------------------Functions---------------------------#
Function Test-Existence {
    [CmdletBinding()]
    param(
        $value,
        $errorMsg
    )
    try {
        if ($value) {
            return $true
        }
        else {
            throw $errorMsg
        }
    }
    catch {
        Write-Error  $_
    }
}

Function Ask-User {
    [CmdletBinding()]
    param(
        $prompt
    )
    Write-Host ($prompt + " (Default is Yes)") -ForegroundColor Yellow
    $readHost = Read-Host " ( y / n )"
    Switch ($readHost) {
        Y { return $true } 
        N { return $false } 
        Default { return $true }
    }
}
#-------------------------Script Setup-------------------------#
if (!$useMFA) { $creds = Get-Credential -Message 'Please sign in to your Global Admin account:' -UserName $adminUPN }

Test-Existence((Get-Module AzureAD-Preview), 'The AzureAD Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module AzureADPreview
if ($useMFA) { Connect-AzureAD -ErrorAction Stop }
else { Connect-AzureAD -Credential $creds -ErrorAction Stop }

$groupOwner = $adminUPN
$sharepointMasterSiteNameShort = $sharepointMasterSiteName.replace(" ", "")
$SharePointMasterSiteURL = $sharepointBaseUrl + "sites/" + $sharepointMasterSiteNameShort

Test-Existence((Get-Module MicrosofTeams), 'The MicrosoftTeams Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module MicrosoftTeams
if ($useMFA) { Connect-MicrosoftTeams -ErrorAction Stop }
else { Connect-MicrosoftTeams -Credential $creds -ErrorAction Stop }

Test-Existence((Get-Module SharePointPnPPowerShellOnline), 'The SharePointPnPPowerShellOnline Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue #Always outputs warning due to unapproved verbs

#Import CSV of Teams
$locationsList = Import-Csv -Path $locationsCsvFile

#Import CSV of Channels/Lists
$subLocationsList = Import-Csv -Path $subLocationsCsvFile

#Import JSON formatting file
$jsonContent = Get-Content $spviewJsonFilePath
$jsonContent | ConvertFrom-Json | Out-Null

#Prepare Microsoft Graph API Calls
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
if ($useMFA) { Connect-PnPOnline -Url $SharePointMasterSiteURL -UseWebLogin -ErrorAction Stop }
else { Connect-PnPOnline -Url $SharePointMasterSiteURL -Credential $creds -ErrorAction Stop }
<##
#-----------------Create Teams and add Members-----------------#
foreach ($location in $locationsList) {
    #Create Team with policies
    $teamName = $location.LocationName + " " + $teamNameSuffix 
    $teamShortName = $location.LocationName.replace(' ', '')
    #check for existing Team/Site
    $existingTeamName = Get-Team -DisplayName $teamName
    $existingTeamMail = Get-Team -MailNickName $teamShortName
    if ($existingTeamName -or $existingTeamMail) {
        $userOption = Ask-User("Existing team found for '$teamName' or '$teamShortName'. Would you like to continue and use this existing team?")
        if ($userOption -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
    }
    else {
        Write-Host "Creating '$teamName' team." -ForegroundColor Green
        New-Team -DisplayName $teamName -Visibility Private -Owner $groupOwner -AllowAddRemoveApps $false -AllowCreateUpdateChannels $false -AllowCreateUpdateRemoveConnectors $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -MailNickName $teamShortName -ErrorAction Stop
        #Add members to team
        $teamID = (Get-AzureADGroup -Filter "DisplayName eq '$teamName'").ObjectID
        #If team has not provisioned yet to AzureAD, keep checking every minute
        while (!$teamID) {
            Write-Host "Team '$teamName' needs more time to provision before members can be added. Trying agin in 60 seconds." -ForegroundColor Green
            Start-sleep -Seconds 60
            $teamID = (Get-AzureADGroup -Filter "DisplayName eq '$teamName'").ObjectID
        }
        $groupName = $location.MembersGroupName
        if ($groupName -ne "") {
            $groupID = (Get-AzureADGroup -Filter "DisplayName eq '$groupName'").ObjectID
            #Ask for new Group Name if none or more than one were found
            while (!($groupID) -or ($groupID.count -gt 1)) {
                $userOption = Ask-User("Unable to find Azure AD Group with the exact namne of '$groupName', or found multiple. Would you like to cancel or specify a new group name?")
                if ($userOption -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
                else {
                    $groupName = Read-Host "Exact Group Name of users to be added to '$teamName' Team:"
                    $groupID = (Get-AzureADGroup -Filter "DisplayName eq '$groupName'").ObjectID
                }
            }
            $groupMembers = Get-AzureADGroupMember -ObjectId $groupID
            Write-Host "Adding Group Membership to '$teamName' team." -ForegroundColor Green
            foreach ($member in $groupMembers) {
                Add-AzureADGroupMember -ObjectId $teamID -RefObjectId $member.ObjectID -ErrorAction SilentlyContinue #silentlycontinue if user is already in the Team
            }
        }
    }
}
#>
foreach ($sublocation in $sublocationsList) {
    #--------------Add Views to List----------------#
    $list = "Lists/" + $sharepointMasterListName.replace(' ', '')
    $locationName = $sublocation.LocationName
    $teamName = $sublocation.LocationName + " " + $teamNameSuffix 
    $teamID = (Get-AzureADGroup -Filter "DisplayName eq '$teamName'").ObjectID
    $teamShortName = $locationName.replace(' ', '')
    $subLocationName = $sublocation.LocationSubName
    $subShortName = $sublocationName.replace(' ', '')
    $viewName = $teamShortName + "-" + $subShortName
    $viewQuery = "<OrderBy><FieldRef Name='ID' /></OrderBy><Where><And><Eq><FieldRef Name='RoomLocation'/><Value Type='Text'>$locationName</Value></Eq><Eq><FieldRef Name='RoomSubLocation' /><Value Type='Text'>$subLocationName</Value></Eq></And></Where>"
    Write-Host "Adding '$viewName' View to Master SharePoint List." -ForegroundColor Green
    Add-PnPView -List $list -Title $viewName -Fields Title, RoomLocation, RoomSubLocation, MeetingLink -Query $viewQuery -ErrorAction SilentlyContinue | Out-Null #Bug in PnP cmdlet, so siletlycontinue required
    Add-PnPView -List $list -Title ($viewName + "-Share") -Fields Title, RoomLocation, "Share Externally", "Reset Room", SharedWith, LastReset -Query $viewQuery -ErrorAction SilentlyContinue | Out-Null #Bug in PnP cmdlet, so siletlycontinue required
    Write-Host "Pausing for 20 seconds for provisioning." -ForegroundColor Green
    Start-Sleep -Seconds 20
    $view = $null
    while (!$view) {
        try {
            $view = Get-PnPView -List $list -Identity $viewName
        }
        catch {
            Write-Host "The view is not ready yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
            Start-Sleep 60
        }
    }
    $view.CustomFormatter = $jsonContent
    $view.Update()
    $view.Context.ExecuteQuery()#
    $viewUrl = ($SharePointMasterSiteURL + "/" + $list + "/" + $viewName + ".aspx").Replace("-","")
    $viewUrl2 = ($SharePointMasterSiteURL + "/" + $list + "/" + ($viewName + "-Share") + ".aspx").Replace("-","")
    $viewUrlEncoded = [System.Web.HTTPUtility]::UrlEncode($viewUrl)
    $viewUrlEncoded2 = [System.Web.HTTPUtility]::UrlEncode($viewUrl2)
    $viewUrl = $viewUrl.replace(" ", "%20") #Needs to be after the encoding step otherwise encoding will encode the '%' symbol
    $viewUrl2 = $viewUrl2.replace(" ", "%20") #Needs to be after the encoding step otherwise encoding will encode the '%' symbol
    Write-Host "Disconnecting SharePoint Online" -ForegroundColor Green
   
    #-------------------Create Channels------------------------#
    #Create Channel
    Write-Host "Creating channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $channelApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels")
    $channelBody = @"
            {"displayName": "$sublocationName", "isFavoriteByDefault": true}
"@
    $newChannel = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $channelApiUrl -Body $channelBody -Method Post -ContentType 'application/json'
    
    #Add SPO List as Tab
    Write-Host "Adding '$tabName' tab to channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs")
    $tabBody = @"
        {
            "displayName": "$tabName",
            "teamsApp@odata.bind" : "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88",
            "configuration": {
              "entityId": "",
              "contentUrl": "$SharePointMasterSiteURL/_layouts/15/teamslogon.aspx?spfx=true&dest=$viewUrlEncoded",
              "websiteUrl": "$viewUrl",
              "removeUrl": null
            }
        }
"@
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $tabApiUrl -Body $tabBody -Method Post -ContentType 'application/json'

    #Add SPO List as Tab
    Write-Host "Adding '$sharetabName' tab to channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs")
    $tabBody = @"
            {
                "displayName": "$sharetabName",
                "teamsApp@odata.bind" : "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88",
                "configuration": {
                  "entityId": "",
                  "contentUrl": "$SharePointMasterSiteURL/_layouts/15/teamslogon.aspx?spfx=true&dest=$viewUrlEncoded2",
                  "websiteUrl": "$viewUrl2",
                  "removeUrl": null
                }
            }
"@
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $tabApiUrl -Body $tabBody -Method Post -ContentType 'application/json'
    
    
    #Get Wiki Tab
    Write-Host "Removing 'Wiki' tab from channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $tabs = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $tabApiUrl -Method GET -ContentType 'application/json'
    $wikiID = ($tabs.value | Where-Object { $_.name -eq "Wiki" }).id
    #Delete Wiki Tab
    $wikiApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs/" + $wikiID)
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $wikiApiUrl -Method DELETE -ContentType 'application/json'
}

Disconnect-PnPOnline

Write-Host "Script Complete." -ForegroundColor Green