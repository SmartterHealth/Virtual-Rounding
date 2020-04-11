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
Please see https://github.com/justinkobel/Virtual-Rounding/
#>

$configFilePath = "C:\Users\justink\GitHub\Virtual-Rounding\Scripts\RunningConfig.json"

$configFile = Get-Content -Path $configFilePath | ConvertFrom-Json


$teamNameSuffix = $configFile.GroupConfiguration.RoundingTeamSuffix
$clientId = $configFile.ClientCredential.Id
$clientSecret = $configFile.ClientCredential.Secret
$tenantName = $configFile.TenantInfo.TenantName
$tenantId = $configFile.TenantInfo.TenantId
$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl
$botCallbackUrl = $callbackUrl = "https://prod-77.eastus.logic.azure.com:443/workflows/1650bb4ce5624c578d1dda634f945957/triggers/manual/paths/invoke"
$useMFA = $configFile.TenantInfo.MFARequired
$adminUPN = $configFile.TenantInfo.GlobalAdminUPN

#--------------------------Functions---------------------------#
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

function Out-BoundView{
    Param(
        [Parameter(ValueFromPipeline = $true)]$viewPath,
        $model
    )
    $view = [String]::Join("",(Get-Content -Path $viewPath))
    $properties = $model.PsObject.Properties | Select-Object -ExpandProperty Name
    foreach($property in $properties){
        $v = $model.($property)
        $view = $view.Replace("{{$($property)}}", $v)
    }
    return $view
}

#------- Script Setup -------#

if($null -eq $credentials)
{
    $Credentials = Get-Credential
}

if($useMFA) {
    Connect-AzureAD 
}
else {
    Connect-AzureAD -Credential $Credentials
}


$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $clientID
    Client_Secret = $clientSecret
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

#temporary, per site:
$sharepointSiteUrl = "$sharepointBaseUrl/sites/BHMCWomenandChildren"

Connect-PnPOnline -Url $sharepointSiteUrl -UseWebLogin

Get-PnpList -Includes ContentTypes,ItemCount | foreach-object {
    $list = $_

    if($list.ContentTypes | Where-Object Name -eq "VirtualRoundingRoom")
    {
        $listItems = Get-PnPListItem -List $list
        $listItems | ForEach-Object  {
            $roundingListItem = $_

            $urlOfMeetingToJoin = $roundingListItem["MeetingLink"].Description
            
            $decodeUrl = [System.Web.HttpUtility]::UrlDecode($urlOfMeetingToJoin)

            $splitUrl = $decodeUrl.Split("/")

            $threadId = $splitUrl[5]
            $contextSplit = $splitUrl[6].Split(":")
            $organizerId = $contextSplit[2]
            $organizerId = $organizerId.Substring(1, $organizerId.Length -2).Trimend('"')

            $teamsMeeting = New-Object PSObject -property @{
                callbackUrl = "https://www.kizan.com/"
                tenantId = $tenantId
                organizerId = $organizerId
                threadId = $threadId
            }

            $body = "C:\Users\justink\GitHub\Virtual-Rounding\Scripts\JoinBotRequestBody.json" | out-boundView -model $teamsMeeting


            Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/communications/calls" -Body $body -Headers @{Authorization = "Bearer $($Tokenresponse.access_token)" } -ContentType "application/json" -Method POST

        }
    }
}
