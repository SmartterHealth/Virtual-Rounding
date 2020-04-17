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
$configFilePath = ".\GitHub\Virtual-Rounding\v2\Scripts\RunningConfig.json"
$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json

$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl
$sharepointMasterSiteName = $configFile.TenantInfo.SPOMasterSiteName
$sharepointMasterListName = $configFile.TenantInfo.SPOMasterListName

$useMFA = $configFile.TenantInfo.MFARequired
$adminUPN = $configFile.TenantInfo.GlobalAdminUPN

#--------------------------Functions---------------------------#
Function Check-Module {
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

Check-Module((Get-Module SharePointPnPPowerShellOnline), 'The SharePointPnPPowerShellOnline Module is not installed. Please see https://aka.ms/virtualroundingcode for more details.') -ErrorAction Stop
Import-Module SharePointPnPPowerShellOnline

#-----------------Create Site and List-----------------#
Write-Host "Connecting to SharePoint Online" -ForegroundColor Green
if ($useMFA) { Connect-PnPOnline -Url $sharepointBaseUrl -UseWebLogin }
else { Connect-PnPOnline -Url $sharepointBaseUrl -Credential $creds }

$existingSite = Get-PnPSiteSearchQueryResults -Query "Title:$sharepointMasterSiteName"
if ($existingSite) { $existingSiteTrue = $true }
while ($existingSiteTrue -eq $true) {
    $userOption = Ask-User("An existing site already exists with the name of '$sharepointMasterSiteName'. Would you like to cancel or specify a new site name?")
    if ($userOption -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
    else {
        $sharepointMasterSiteName = Read-Host "New Site Name:"
        $existingSite = Get-PnPSiteSearchQueryResults -Query "Title:$sharepointMasterSiteName"
        if (!$existingSite) { $existingSiteTrue = $false }
    }
}
$sharepointMasterSiteNameShort = $sharepointMasterSiteName.replace(" ", "")
$SharePointMasterSiteURL = $sharepointBaseUrl + "sites/" + $sharepointMasterSiteNameShort

Write-Host "Creating Site '$sharepointMasterSiteName'" -ForegroundColor Green
New-PnPTenantSite -Title $sharepointMasterSiteName -Url $SharePointMasterSiteURL -Owner $adminUPN -TimeZone 11 -ErrorAction Stop

Write-Host "Connecting to Site '$sharepointMasterSiteName'" -ForegroundColor Green
$siteReady = $false
while ($siteReady -eq $false) {
    try {
        if ($useMFA) { Connect-PnPOnline -Url $SharePointMasterSiteURL -UseWebLogin }
        else { Connect-PnPOnline -Url $SharePointMasterSiteURL -Credential $creds }
        $siteReady = $true
    }
    catch {
        $userOption2 = Ask-User("Unable to connect to SharePoint Site. This is commonly a result of a provisioning delay. Would you like to have the script pause for 5 minutes and try again?")
        if ($userOption2 -eq $false) { Write-Host "Script stopping by user request" -ForegroundColor Red -ErrorAction Stop }
        if ($userOption2 -eq $true) { Write-Host "Pausing for 5 minutes to wait for SharePoint Site Provisioning." -ForegroundColor Green; Start-Sleep -Seconds 300 }
    }
}

Write-Host "Setting up Site Columns & Content Type" -ForegroundColor Green
Add-PnPField -Type Text -InternalName "RoomLocation" -DisplayName "Room Location" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "RoomSubLocation" -DisplayName "Room SubLocation" -Group "VirtualRounding"
Add-PnPField -Type URL -InternalName "MeetingLink" -DisplayName "Meeting Link" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "EventID" -DisplayName "EventID" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "Share Externally" -DisplayName "Share Externally" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "Reset Room" -DisplayName "Reset Room" -Group "VirtualRounding"
Add-PnPField -Type DateTime -InternalName "LastReset" -DisplayName "Last Reset" -Group "VirtualRounding"
Add-PnPField -Type Number -InternalName "SharedWith" -DisplayName "Shared With" -Group "VirtualRounding"
Add-PnPField -Type DateTime -InternalName "LastShare" -DisplayName "Last Share" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "RoomUPN" -DisplayName "Room UPN" -Group "VirtualRounding"
Add-PnPField -Type Text -InternalName "Patient Name" -DisplayName "Patient Name" -Group "VirtualRounding"
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
Add-PnPFieldToContentType -Field "EventID" -ContentType $contentType
Add-PnPFieldToContentType -Field "Share Externally" -ContentType $contentType
Add-PnPFieldToContentType -Field "Reset Room" -ContentType $contentType
Add-PnPFieldToContentType -Field "LastReset" -ContentType $contentType
Add-PnPFieldToContentType -Field "SharedWith" -ContentType $contentType
Add-PnPFieldToContentType -Field "LastShare" -ContentType $contentType
Add-PnPFieldToContentType -Field "RoomUPN" -ContentType $contentType
Add-PnPFieldToContentType -Field "Patient Name" -ContentType $contentType

Write-Host "Creating SharePoint List '$sharepointMasterListName'." -ForegroundColor Green
$listShortName = $sharepointMasterListName.replace(" ","")
$listUrl = "Lists/$listShortName"
New-PnPList -Title $sharepointMasterListName -Url $listUrl -Template GenericList
Start-Sleep -Seconds 5
while (!$list) {
    try {
        $list = Get-PnPList -Identity $listUrl
    }
    catch {
        Write-Host "List is not provisioned yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
        Start-Sleep 60
    }
}

Write-Host "Enabling Content Types for SharePoint List '$sharepointMasterListName'." -ForegroundColor Green
Set-PnPList -Identity $list -EnableContentTypes $true
Write-Host "Adding 'VirtualRoundingRoom' Content Type to SharePoint List '$sharepointMasterListName'." -ForegroundColor Green
Add-PnPContentTypeToList -List $list -ContentType $contentType -DefaultContentType -ErrorAction SilentlyContinue | Out-Null #Bug in PnP cmdlet, so SilentlyContinue required
$newContentType = $null
while (!$newContentType) {
    try {
        $newContentType = Get-PnPContentType -List $list | Where-Object { $_.Name -eq "VirtualRoundingRoom" }
    }
    catch {
        Write-Host "Content Type is not added to list yet. Waiting 1 minute for provisioning. This will repeat each minute until ready." -ForegroundColor Yellow
        Start-Sleep 60
    }
}

Disconnect-PnPOnline

Write-Host "Script Complete." -ForegroundColor Green