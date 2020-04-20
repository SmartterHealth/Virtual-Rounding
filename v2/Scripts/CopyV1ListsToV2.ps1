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
$configFilePath = "C:\Users\justink\OneDrive - KiZAN Technologies, LLC\Scripts\BrowardV2\Scripts\RunningConfig.json"
$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json

$clientID = $configFile.ClientCredential.Id
$clientSecret = $configFile.ClientCredential.Secret
$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl
$sharepointMasterSiteName = $configFile.TenantInfo.SPOMasterSiteName
$sharepointMasterListName = $configFile.TenantInfo.SPOMasterListName
$locationsCsvPath = $configFile.LocationCsvPaths.Locations
$TenantName = $configFile.TenantInfo.TenantName
$roomUpnSuffix = $configFile.TenantInfo.RoomUPNSuffix
$spviewJsonFilePath = $configFile.ViewJson.SPViewJsonFilePath
$tabName = $configFile.TeamsInfo.tabName
$shareTabName = $configFile.TeamsInfo.shareTabName

$useMFA = $configFile.TenantInfo.MFARequired

#------- Script Setup -------#

$sharepointMasterSiteUrl = $sharepointBaseUrl + "sites/" + $sharepointMasterSiteName

if (!$useMFA -and $null -ne $creds) { $creds = Get-Credential -Message 'Please sign in to your Global Admin account:' -UserName $adminUPN }

$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

$locationsList = Import-Csv -Path $locationsCsvPath

#Import JSON formatting file for v1-style SharePoint views
$viewFormattingJson = Get-Content $spviewJsonFilePath
$viewFormattingJson | ConvertFrom-Json | Out-Null

Connect-MicrosoftTeams -Credential $creds

Write-Host "Connecting to SharePoint Online Master Site" -ForegroundColor Green
if ($useMFA) { $masterConnection = Connect-PnPOnline -Url $sharepointMasterSiteUrl -UseWebLogin -ReturnConnection }
else { $masterConnection = Connect-PnPOnline -Url $sharepointMasterSiteUrl -Credential $creds -ReturnConnection }

$masterList = Get-PnPList -Identity $sharepointMasterListName -Connection $masterConnection
$masterCtype = Get-PnPContentType -List $masterList -Connection $masterConnection -Identity "VirtualRoundingRoom"

foreach ($location in $locationsList) {

    $teamName = $location.LocationName + " " + $teamNameSuffix 
    $teamShortName = $location.LocationName.replace(' ', '')
    
    #check for existing Team/Site
    $existingTeam = Get-Team -MailNickName $teamShortName
    if ($existingTeam) {
        $sharepointSiteUrl = "$sharepointBaseUrl/sites/$teamShortName"

        $locationTeamChannels = Get-TeamChannel -GroupId $existingTeam.GroupId

        Write-Host "Connecting to SharePoint Online v1 site $location" -ForegroundColor Green
        if ($useMFA) { $locationConnection = Connect-PnPOnline -Url $sharepointSiteUrl -UseWebLogin -ReturnConnection }
        else { $locationConnection = Connect-PnPOnline -Url $sharepointSiteUrl -Credential $creds -ReturnConnection }

        Get-PnpList -Includes ContentTypes, ItemCount -Connection $locationConnection | foreach-object {
        
            $list = $_

            if ($list.ContentTypes | Where-Object Name -eq "VirtualRoundingRoom") {
                Write-Host "Preparing to copy for " $list.Title
                $listItems = Get-PnPListItem -List $list -Connection $locationConnection
                
                $listItems | ForEach-Object {
                    
                    $v1RoundingListItem = $_
                    $Title = $v1RoundingListItem["Title"]
                    $EventID = $v1RoundingListItem["EventID"]
                    $lastReset = Get-Date
                    $lastShare = Get-Date
                    $meetingLink = $v1RoundingListItem["MeetingLink"].Url
                    $roomLocation = $v1RoundingListItem["RoomLocation"]
                    $roomSubLocation = $v1RoundingListItem["RoomSubLocation"]

                    $roomUpn = "$Title$roomUpnSuffix"
                    
                    Write-Host "pre new item"

                    Add-PnPListItem -List $masterList -Connection $masterConnection -ContentType $masterCtype -Values @{"Title" = $Title; "EventID" = $EventID; "LastReset" = $lastReset; "LastShare" = $lastShare; "MeetingLink" = $meetingLink; "RoomLocation" = $roomLocation; "RoomSubLocation" = $roomSubLocation; "RoomUPN" = $roomUpn } | Out-Null #darn ISE

                }
            
                $locationName = $location.LocationName
                $subLocationName = $list.Title

                $subLocationShortName = $list.Title.Replace(' ', '')
                $viewName = "$teamShortName-$subLocationShortName"
                $viewQuery = "<OrderBy><FieldRef Name='ID' /></OrderBy><Where><And><Eq><FieldRef Name='RoomLocation'/><Value Type='Text'>$locationName</Value></Eq><Eq><FieldRef Name='RoomSubLocation' /><Value Type='Text'>$subLocationName</Value></Eq></And></Where>"

                $existingView = Get-PnPView -List $masterList -Identity $viewName -Connection $masterConnection -ErrorAction SilentlyContinue
                if ($existingView) {
                    Write-Host "Skipping view creation for $viewName since it already exists"
                }
                else {
                    Write-Host "Creating a view for sublocation " $viewName

                    Write-Host "Adding '$viewName' Views to Master SharePoint List." -ForegroundColor Green
                    Add-PnPView -List $masterList -Title $viewName -Fields Title, RoomLocation, RoomSubLocation, MeetingLink -Query $viewQuery -Connection $masterConnection | Out-Null #Bug in PnP cmdlet, so siletlycontinue required
                    Add-PnPView -List $masterList -Title ($viewName + "-Share") -Fields Title, RoomLocation, "Share Externally", "Reset Room", SharedWith, LastReset -Query $viewQuery -Connection $masterConnection | Out-Null #Bug in PnP cmdlet, so siletlycontinue required
                    Write-Host "Pausing for 20 seconds for provisioning." -ForegroundColor Green
                    Start-Sleep -Seconds 20
                    $view = $null
                    while (!$view) {
                        try {
                            $view = Get-PnPView -List $masterList -Identity $viewName -Connection $masterConnection
                        }
                        catch {
                            Write-Host "The view is not ready yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
                            Start-Sleep 60
                        }
                    }
                    $view.CustomFormatter = $viewFormattingJson
                    $view.Update()
                    $view.Context.ExecuteQuery()
                    $viewUrl = ($SharePointMasterSiteURL + "/" + $list + "/" + $viewName + ".aspx").Replace("-", "")
                    $viewUrl2 = ($SharePointMasterSiteURL + "/" + $list + "/" + ($viewName + "-Share") + ".aspx").Replace("-", "")
                    $viewUrlEncoded = [System.Web.HTTPUtility]::UrlEncode($viewUrl)
                    $viewUrlEncoded2 = [System.Web.HTTPUtility]::UrlEncode($viewUrl2)
                    $viewUrl = $viewUrl.replace(" ", "%20") #Needs to be after the encoding step otherwise encoding will encode the '%' symbol
                    $viewUrl2 = $viewUrl2.replace(" ", "%20") #Needs to be after the encoding step otherwise encoding will encode the '%' symbol
                }

                $subLocationChannel = $locationTeamChannels | Where-Object { $_.DisplayName -eq $subLocationName }
            
                $teamID = $existingTeam.GroupId

                #Add SPO List as Tab
                Write-Host "Adding '$tabName' tab to channel '$sublocationName' in the '$teamName' Team." -ForegroundColor Green
                $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $subLocationChannel.id + "/tabs")
                $tabBody = @"
        {
            "displayName": "$tabName v2",
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
                $tabApiUrl = ("https://graph.microsoft.com/beta/teams/" + $teamID + "/channels/" + $subLocationChannel.id + "/tabs")
                $tabBody = @"
            {
                "displayName": "$sharetabName v2",
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
            }

            Disconnect-PnPOnline -Connection $locationConnection
        }
    }
    else {
        Write-Host "Unable to migration for location $teamShortName" -ForegroundColor Red
    }

}