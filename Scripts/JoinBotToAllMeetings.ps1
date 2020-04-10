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

$teamNameSuffix = $configFile.GroupConfiguration.RoundingTeamSuffix
$clientId = $configFile.ClientCredential.Id
$clientSecret = $configFile.ClientCredential.Secret
$tenantName = $configFile.TenantInfo.TenantName
$sharepointBaseUrl = $configFile.TenantInfo.SPOBaseUrl

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

#------- Script Setup -------#