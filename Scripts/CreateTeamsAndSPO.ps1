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
Please see https://github.com/SmartterHealth/Virtual-Rounding/
#>

#--------------------------Variables---------------------------#
$configFilePath = ".\Scripts\RunningConfig.json"


$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json


$locationsCsvFile = $configFile.LocationCsvPaths.Locations
$subLocationsCsvFile = $configFile.LocationCsvPaths.SubLocations
$groupOwner = $configFile.TenantInfo.MeetingSchedulingUser
$spviewJsonFilePath = $configFile.ViewJson.SPViewJsonFilePath
$teamNameSuffix = $configFile.GroupConfiguration.RoundingTeamPrefix
$clientId = $configFile.ClientCredential.Id
$clientSecret = $configFile.ClientCredential.Secret
$tenantName = $configFile.TenantInfo.TenantName
$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl

#-------------------------Script Setup-------------------------#
$credentials = Get-Credential

Import-Module AzureAD
Connect-AzureAD -Credential $credentials

Import-Module MicrosoftTeams
Connect-MicrosoftTeams -Credential $credentials

Import-Module SharePointPnPPowerShellOnline

#Import CSV of Teams
$locationsList = Import-Csv -Path $locationsCsvFile

#Import CSV of Channels
$subLocationsList = Import-Csv -Path $subLocationsCsvFile

#Import JSON formatting file
$spViewJson = Get-Content $spviewJsonFilePath
$spViewJson | ConvertFrom-Json | Out-Null

#Prepare Microsoft Graph API Calls
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

#-----------------Create Teams and add Members-----------------#
foreach ($location in $locationsList) {
    #Create Team with policies
    $teamName = $location.LocationName + " " + $teamNameSuffix 
    $teamShortName = $location.LocationName.replace(' ','')
    New-Team -DisplayName $teamName -Visibility Private -Owner $groupOwner -AllowAddRemoveApps $false -AllowCreateUpdateChannels $false -AllowCreateUpdateRemoveConnectors $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -MailNickName $teamShortName
    #Add members to team
    $groupID = (Get-AzureADGroup -SearchString $location.MembersGroupName).ObjectID
    $groupMembers = Get-AzureADGroupMember -ObjectId $groupID
    foreach ($member in $groupMembers) {
        Add-AzureADGroupMember -ObjectId $teamID -RefObjectId $member.ObjectID
    }
    $teamSpoUrl = $sharepointBaseUrl + "sites/" + $teamShortName
    Connect-PnPOnline -Url $teamSpoUrl -Credentials $credentials
    #--------------------Setup Site Colums---------------------#
    Add-PnPField -Type Text -InternalName "RoomLocation" -DisplayName "Room Location" -Group "VirtualRounding"
    Add-PnPField -Type Text -InternalName "RoomSubLocation" -DisplayName "Room SubLocation" -Group "VirtualRounding"
    Add-PnPField -Type URL -InternalName "MeetingLink" -DisplayName "Meeting Link" -Group "VirtualRounding"
    Add-PnPField -Type Text -InternalName "EventID" -DisplayName "EventID" -Group "VirtualRounding"
    Add-PnPField -Type User -InternalName "RoomAccount" -DisplayName "RoomAccount" -Group "VirtualRounding"
    Add-PnPContentType -Name "VirtualRoundingRoom" -Group "VirtualRounding"
    Start-Sleep -Seconds 5
    $contentType = Get-PnPContentType -Identity "VirtualRoundingRoom"
    Add-PnPFieldToContentType -Field "RoomLocation" -ContentType $contentType
    Add-PnPFieldToContentType -Field "RoomSubLocation" -ContentType $contentType
    Add-PnPFieldToContentType -Field "MeetingLink" -ContentType $contentType
    Add-PnPFieldToContentType -Field "EventID" -ContentType $contentType
    Add-PnPFieldToContentType -Field "RoomAccount" -ContentType $contentType
    Disconnect-PnPOnline
}

foreach ($sublocation in $sublocationsList) {
    #--------------Create Lists and Add Columns----------------#
    $sublocationName = $sublocation.LocationSubName
    $teamName = $sublocation.LocationName + " " + $teamNameSuffix
    $teamID = (Get-AzureADGroup -SearchString $teamName).ObjectID
    $teamShortName = $sublocation.LocationName.replace(' ','')
    $teamSpoUrl = $sharepointBaseUrl + "sites/" + $teamShortName
    Connect-PnPOnline -Url $teamSpoUrl -Credentials $credentials
    $contentType = Get-PnPContentType -Identity "VirtualRoundingRoom"

    New-PnPList -Title $sublocationName -Template GenericList
    Start-Sleep -Seconds 5
    $sublocationShortName = $sublocationName.replace('-','')
    $list = Get-PnPList -Identity ("Lists/" + $sublocationShortName)
    Set-PnPList -Identity $sublocationShortName -EnableContentTypes $true
    #set Permissions?
    Add-PnPContentTypeToList -List $list -ContentType $contentType -DefaultContentType
    $newView = Add-PnPView -List $list -Title Meetings -SetAsDefault -Fields Title, RoomLocation, MeetingLink
    Start-Sleep -Seconds 20 #toolong?
    $view = Get-PnPView -List $list -Identity Meetings
    $view.CustomFormatter = $spViewJson
    $view.Update()
    $view.Context.ExecuteQuery()
    $viewUrl = ($teamSpoUrl + "/Lists/" + $sublocationShortName + "/Meetings.aspx")
    $viewUrlEncoded = [System.Web.HTTPUtility]::UrlEncode($viewUrl)
    $viewUrl = $viewUrl.replace(" ","%20")
    #-------------------Create Channels------------------------#
    #Create Channel
    $channelApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels")
    $channelBody = @"
            {"displayName": "$sublocationName"}
"@
    $newChannel = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $channelApiUrl -Body $channelBody -Method Post -ContentType 'application/json'
    #Add SPO List as Tab
    $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs")
    $tabBody = @"
        {
            "displayName": "Join a Room",
            "teamsApp@odata.bind" : "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88",
            "configuration": {
              "entityId": "sharepointtab_0.8309667588452743",
              "contentUrl": "$teamSpoUrl/_layouts/15/teamslogon.aspx?spfx=true&dest=$viewUrlEncoded",
              "websiteUrl": "$viewUrl",
              "removeUrl": null
            }
        }
"@
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $tabApiUrl -Body $tabBody -Method Post -ContentType 'application/json'
    #Get Wiki Tab
    $tabs = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $tabApiUrl -Method GET -ContentType 'application/json'
    $wikiID = ($tabs.value | ? { $_.name -eq "Wiki" }).id
    #Delete Wiki Tab
    $wikiApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $newChannel.id + "/tabs/" + $wikiID)
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -Uri $wikiApiUrl -Method DELETE -ContentType 'application/json'
    
    Disconnect-PnPOnline
}
