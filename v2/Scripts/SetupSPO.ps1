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
$configFilePath = ".\Scripts\RunningConfig.json"
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
$teamsTabName = $configFile.TeamsInfo.TabnName
$timeZoneNumber = $configFile.TimeZoneInfo.UTCOffset

$useMFA = $configFile.TenantInfo.MFARequired
$adminUPN = $configFile.TenantInfo.GlobalAdminUPN

if(!$useMFa -or !$adminUPN){Write-Host "Missing JSON values"} #FINISHME

#--------------------------Functions---------------------------#
Function Check-Module {
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

Function Ask-User {
    [CmdletBinding()]
    param(
        $prompt
    )
    Write-Host ($prompt + " (Default is Yes)") -ForegroundColor Yellow
    $readHost = Read-Host " ( y / n )"
    Switch ($readHost){
        Y {return $true} 
        N {return $false} 
        Default {return $true}
    }
}
#-------------------------Script Setup-------------------------#
if (!$useMFA) {$creds = Get-Credential -Message 'Please sign in to your Global Admin account:' -UserName $adminUPN}

Check-Module((Get-Module AzureAD-Preview),'The AzureAD Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module AzureAD
if ($useMFA) {Connect-AzureAD -ErrorAction Stop}
else {Connect-AzureAD -Credential $creds -ErrorAction Stop}

$groupOwner = $adminUPN
$sharepointMasterSiteNameShort = $sharepointMasterSiteName.replace(" ","")

Check-Module((Get-Module SharePointPnPPowerShellOnline),'The SharePointPnPPowerShellOnline Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module SharePointPnPPowerShellOnline

#Import JSON formatting file
$jsonContent = Get-Content $spviewJsonFilePath
$jsonContent | ConvertFrom-Json | Out-Null

#-----------------Create Teams and add Members-----------------#
Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
if ($useMFA) { Connect-PnPOnline -Url $sharepointBaseUrl -UseWebLogin }
else { Connect-PnPOnline -Url $sharepointBaseUrl -Credential $creds }

$existingSite = Get-PnPSiteSearchQueryResults -Query "Title:$sharepointSiteMasterSiteName"
if ($existingSite){$existingSiteTrue = $true}
while ($existingSiteTrue -eq $true){
    $userOption = Ask-User("An existing site already exists with the name of '$sharepointSiteMasterSiteName'. Would you like to cancel or specify a new site name?")
    if ($userOption -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
    else {
        $groupName = Read-Host "New Site Name:"
        $existingSite = Get-PnPSiteSearchQueryResults -Query "Title:$sharepointSiteMasterSiteName"
        if (!$existingSite){$existingSiteTrue = $false}
    }
}

Write-Host "Creating Site '$sharepointMasterSiteName'" -ForegroundColor Green
New-SPOTenantSite 

Write-Host "Connecting to Site '$sharepointMasterSiteName'" -ForegroundColor Green
while ($siteReady -eq $false) {
    try {
        if ($useMFA) { Connect-PnPOnline -Url $teamSpoUrl -UseWebLogin }
        else { Connect-PnPOnline -Url $teamSpoUrl -Credential $creds }
        $siteReady = $true
    }
    catch {
        $userOption2 = Ask-User("Unable to connect to SharePoint Site. This is commonly a result of a provisioning delay. Would you like to have the script pause for 5 minutes and try again?")
        if ($userOption2 -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
        if ($userOption2 -eq $true) { Write-Host "Pausing for 5 minutes to wait for SharePoint Site Provisioning." -ForegroundColor Green; Start-Sleep -Seconds 300 }
    }
}











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
            Start-sleep -Seconds 60
            $teamID = (Get-AzureADGroup -Filter "DisplayName eq '$teamName'").ObjectID
        }
        $groupName = $location.MembersGroupName
        $groupID = (Get-AzureADGroup -Filter "DisplayName eq '$groupName'").ObjectID
        $groupMembers = Get-AzureADGroupMember -ObjectId $groupID
        Write-Host "Adding Group Membership to '$teamName' team." -ForegroundColor Green
        foreach ($member in $groupMembers) {
            Add-AzureADGroupMember -ObjectId $teamID -RefObjectId $member.ObjectID
        }
    
        $teamSpoUrl = $sharepointBaseUrl + "sites/" + $teamShortName
        $siteReady = $false
        Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
        while ($siteReady -eq $false) {
            try {
                if ($useMFA) { Connect-PnPOnline -Url $teamSpoUrl -UseWebLogin }
                else { Connect-PnPOnline -Url $teamSpoUrl -Credential $creds }
                $siteReady = $true
            }
            catch {
                $userOption2 = Ask-User("Unable to connect to SharePoint Site. This is commonly a result of a provisioning delay. Would you like to have the script pause for 5 minutes and try again?")
                if ($userOption2 -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
                if ($userOption2 -eq $true) { Write-Host "Pausing for 5 minutes to wait for SharePoint Site Provisioning." -ForegroundColor Green; Start-Sleep -Seconds 300 }
            }
        }
        #--------------------Setup Site Colums---------------------#
        Write-Host "Setting up Site Columns & Content Type" -ForegroundColor Green
        Add-PnPField -Type Text -InternalName "RoomLocation" -DisplayName "Room Location" -Group "VirtualRounding"
        Add-PnPField -Type Text -InternalName "RoomSubLocation" -DisplayName "Room SubLocation" -Group "VirtualRounding"
        Add-PnPField -Type URL -InternalName "MeetingLink" -DisplayName "Meeting Link" -Group "VirtualRounding"
        Add-PnPField -Type Text -InternalName "EventID" -DisplayName "EventID" -Group "VirtualRounding"
        Add-PnPField -Type Text -InternalName "Share Externally" -DisplayName "Share Externally" -Group "VirtualRounding"
        Add-PnPField -Type Text -InternalName "Reset Room" -DisplayName "Reset Room" -Group "VirtualRounding"
        Add-PnPField -Type DateTime -InternalName "LastReset" -DisplayName "Last Reset" -Group "VirtualRounding"
        Add-PnPField -Type Number -InternalName "SharedWith" -DisplayName "Shared With" -Group "VirtualRounding"
        #Set-PnPDefaultColumnValues
        Add-PnPField -Type DateTime -InternalName "LastShare" -DisplayName "Last Share" -Group "VirtualRounding"
        Add-PnPField -Type Text -InternalName "RoomUPN" -DisplayName "Room UPN" -Group "VirtualRounding"
        Add-PnPContentType -Name "VirtualRoundingRoom" -Group "VirtualRounding" | Out-Null
        Start-Sleep -Seconds 5
        $contentType = $null
        while (!$contentType) {
            try {
                $contentType = Get-PnPContentType -Identity "VirtualRoundingRoom"
            }
            catch {
                Write-Host "Content Type is not provisioned yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
                Start-Sleep 60
            }
        }
        Add-PnPFieldToContentType -Field "RoomLocation" -ContentType $contentType
        Add-PnPFieldToContentType -Field "RoomSubLocation" -ContentType $contentType
        Add-PnPFieldToContentType -Field "MeetingLink" -ContentType $contentType
        Add-PnPFieldToContentType -Field "EventID" -ContentType $contentType -Hidden $true
        Add-PnPFieldToContentType -Field "SharetoFamily" -ContentType $contentType
        Add-PnPFieldToContentType -Field "ResetMeeting" -ContentType $contentType
        Add-PnPFieldToContentType -Field "LastReset" -ContentType $contentType
        Add-PnPFieldToContentType -Field "FamilyInvited" -ContentType $contentType
        Add-PnPFieldToContentType -Field "LastFamilyInvite" -ContentType $contentType
        Add-PnPFieldToContentType -Field "RoomUPN" -ContentType $contentType -Hidden $true
        Disconnect-PnPOnline
    }
}

foreach ($sublocation in $sublocationsList) {
    #--------------Create Lists and Add Columns----------------#
    $sublocationName = $sublocation.LocationSubName
    $sublocationShortName = $sublocationName.replace('-','')
    $teamName = $sublocation.LocationName + " " + $teamNameSuffix
    $teamID = (Get-AzureADGroup -Filter "DisplayName eq '$teamName'").ObjectID
    $teamShortName = $sublocation.LocationName.replace(' ','')
    $teamSpoUrl = $sharepointBaseUrl + "sites/" + $teamShortName
    
    Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
    if ($useMFA) { Connect-PnPOnline -Url $teamSpoUrl -UseWebLogin }
    else { Connect-PnPOnline -Url $teamSpoUrl -Credential $creds }

    $contentType = Get-PnPContentType -Identity "VirtualRoundingRoom"

    $list = Get-PnPList -Identity ("Lists/" + $sublocationShortName)
    if ($list) {
        $userOption = Ask-User("An existing list for '$sublocationName' already exists. Would you like to use this existing list?")
        if ($userOption -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
    }
    else {
        Write-Host "Creating SharePoint List '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
        New-PnPList -Title $sublocationName -Url "Lists/$sublocationshortName" -Template GenericList
        Start-Sleep -Seconds 5
        while(!$list){
            try {
                $list = Get-PnPList -Identity ("Lists/" + $sublocationShortName)
            }
            catch {
                Write-Host "List is not provisioned yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
                Start-Sleep 60
            }
        }
    }
    Write-Host "Enabling Content Types for SharePoint List '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    Set-PnPList -Identity $sublocationShortName -EnableContentTypes $true
    Write-Host "Adding 'VirtualRoundingRoom' Content Type to SharePoint List '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    Add-PnPContentTypeToList -List $list -ContentType $contentType -DefaultContentType -ErrorAction SilentlyContinue | Out-Null #Bug in PnP cmdlet, so SilentlyContinue required
    $newContentType = $null
    while(!$newContentType){
        try {
            $newContentType = Get-PnPContentType -List $list | Where-Object{$_.Name -eq "VirtualRoundingRoom"}
        }
        catch {
            Write-Host "Content Type is not added to list yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
            Start-Sleep 60
        }
    }
    Write-Host "Adding 'Meetings' View to SharePoint List '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    Add-PnPView -List $list -Title Meetings -SetAsDefault -Fields Title, RoomLocation, MeetingLink -ErrorAction SilentlyContinue | Out-Null #Bug in PnP cmdlet, so siletlycontinue required
    Write-Host "Pausing for 20 seconds for provisioning." -ForegroundColor Green
    Start-Sleep -Seconds 20
    $view = $null
    while(!$view){
        try {
            $view = Get-PnPView -List $list -Identity Meetings
        }
        catch {
            Write-Host "The view is not ready yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
            Start-Sleep 60
        }
    }
    $view.CustomFormatter = $jsonContent
    $view.Update()
    $view.Context.ExecuteQuery()
    $viewUrl = ($teamSpoUrl + "/Lists/" + $sublocationShortName + "/Meetings.aspx")
    $viewUrlEncoded = [System.Web.HTTPUtility]::UrlEncode($viewUrl)
    $viewUrl = $viewUrl.replace(" ","%20")
    #####IS THERE A BETTER WAY TO GET THIS VIEW URL?
    Disconnect-PnPOnline
    #-------------------Create Channels------------------------#
    #Create Channel
    Write-Host "Creating channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $channelApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels")
    $channelBody = @"
            {"displayName": "$sublocationName"}
"@
    $newChannel = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $channelApiUrl -Body $channelBody -Method Post -ContentType 'application/json'
    #Add SPO List as Tab (need to make it non-hidden)
    Write-Host "Adding '$teamsTabName' tab to channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
    $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs")
    $tabBody = @"
        {
            "displayName": "$teamsTabName",
            "teamsApp@odata.bind" : "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88",
            "configuration": {
              "entityId": "",
              "contentUrl": "$teamSpoUrl/_layouts/15/teamslogon.aspx?spfx=true&dest=$viewUrlEncoded",
              "websiteUrl": "$viewUrl",
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