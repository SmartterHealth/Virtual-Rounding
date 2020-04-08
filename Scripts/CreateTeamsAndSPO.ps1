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
#Path of the first CSV file. (Columns expected: LocationName, MembersGroupName)
#LOCATION NAMES SHOULD MATCH AccountLocations FROM CreateRooms SCRIPT
$locationsCsvFile = ""
#Path of the first CSV file. (Columns expected: SubLocationName, LocationName)
#SUBLOCATION NAMES SHOULD MATCH AccountSubLocations FROM CreateRooms SCRIPT
$subLocationsCsvFile = ""
#UPNs of A desired Team owner (will apply to all Teams)
#NEEDS TO BE THE SAME ACCOUNT YOU LOG IN WITH DURING THIS SCRIPT
$groupOwner = "Kelly@contosohealthsystem.onmicrosoft.com"
#Path of the JSON file (download from same repository as this script)
$jsonFile = ""
#Team Name Suffix - Text to be added after LocationName to form the team name (use this in the Flow later in the setup process too)
#Example: LocationName:"Building 2" + Suffix:"Patient Rooms" = Team Name: "Building 2 Patient Rooms"
$teamNameSuffix = "Virtual Rounding" #*required*
#Azure AD App Registration Info:
$clientId = "" #from Azure AD App Registration
$tenantName = "contosohealthsystem.onmicrosoft.com" # your onmicrosoft domain
$clientSecret = "" #from Azure AD App Registration
#Define your Tenant SPO Url
$sharepointBaseUrl = "https://contosohealthsystem.sharepoint.com/" #ensure this ends with a /
#-------------------------Script Setup-------------------------#
$credentials = Get-Credential

Import-Module AzureAD
Connect-AzureAD -Credential $credentials

Import-Module MicrosoftTeams
Connect-MicrosoftTeams -Credential $credentials

Import-Module SharePointPnPPowerShellOnline

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
#Import-PSSession $Session -DisableNameChecking

#Import CSV of Teams
$locationsList = Import-Csv -Path $locationsCsvFile

#Import CSV of Channels
$subLocationsList = Import-Csv -Path $subLocationsCsvFile

#Import JSON formatting file
$jsonContent = Get-Content $jsonFile
$jsonContent | ConvertFrom-Json | Out-Null

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
    #check for Team/Site
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
    Add-PnPContentType -Name "VirtualRoundingRoom" -Group "VirtualRounding"
    Start-Sleep -Seconds 5
    $contentType = Get-PnPContentType -Identity "VirtualRoundingRoom"
    Add-PnPFieldToContentType -Field "RoomLocation" -ContentType $contentType
    Add-PnPFieldToContentType -Field "RoomSubLocation" -ContentType $contentType
    Add-PnPFieldToContentType -Field "MeetingLink" -ContentType $contentType
    Add-PnPFieldToContentType -Field "EventID" -ContentType $contentType
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

    #check for list
    New-PnPList -Title $sublocationName -Template GenericList
    Start-Sleep -Seconds 5
    $sublocationShortName = $sublocationName.replace('-','')
    $list = Get-PnPList -Identity ("Lists/" + $sublocationShortName)
    Set-PnPList -Identity $sublocationShortName -EnableContentTypes $true
    #set Permissions?
    #check for content typer
    $newContentType Add-PnPContentTypeToList -List $list -ContentType $contentType -DefaultContentType
    #check for view
    $newView = Add-PnPView -List $list -Title Meetings -SetAsDefault -Fields Title, RoomLocation, MeetingLink
    Start-Sleep -Seconds 20 #toolong?
    $view = Get-PnPView -List $list -Identity Meetings
    $view.CustomFormatter = $jsonContent
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