#Load ConfigMge PoSH Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Branding (change to customer name)
$OrgName = 'MSFT Services'

#Package/Update Share path
#This is the location where the software updates source location will be placed i.e. \\server\share\folder
$PackageShare = "\\SCCMCB\Swcatalog\SoftwareUpdates"

#SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference= 'Continue'
$Error1 = 0

#If running Remotely change these values to the Primary hostanme and FQDN
$SiteServer = $Env:COMPUTERNAME
$SiteServerFQDN = "$SiteServer.$Env:USERDNSDOMAIN"

#Distribution Point 
$DPFQDN = $SiteServerFQDN

#ADR Info
$Collection = "Preview ADR Windows 8 and 8.1"
$ADRName = "$OrgName Preview ADR Windows 8 and 8.1"
$DeployPackageLocation = $PackageShare +"\" + $ADRName
$CMPSSuppressFastNotUsedCheck = $true

#Update Info
$Products = "Windows 8","Windows 8.1"
$UpdateClassifications = "Critical Updates","Definition Updates","Security Updates","Service Packs","Update Rollups","Updates"
$Severity = "Critical","Important","Low","Moderate","None"
 
if (Get-CMDeviceCollection -Name $Collection)
{
    Write-Output "$Collection collection found, continue"
}
Else
{
    Write-Warning "$Collection collection does not exist"
    Break
}
 
# Create Software Update Deployment Package
if (Get-CMSoftwareUpdateDeploymentPackage -Name $ADRName)
{

    Write-Output "$ADRName Software Update Deployment Package found, continue"
}
Else
{
    Write-Warning "$ADRName Software Update Deployment Package does not exist, creating it"
    $NewDeploymentPackage = New-CMSoftwareUpdateDeploymentPackage -Name $ADRName -Path $DeployPackageLocation
}
 
# Distribute the Software Update Deployment Package
Start-CMContentDistribution -DeploymentPackageId $NewDeploymentPackage.PackageID -DistributionPointName $DPFQDN
 
if (Get-CMSoftwareUpdateAutoDeploymentRule -Name $ADRName)
{

    Write-Output "$ADRName Automatic Deployment Rule already exist"
 }
Else
{
    Write-Output "$ADRName Automatic Deployment Rule does not exist, creating it"
     
    $Schedule = New-CMSchedule –RecurInterval Days –RecurCount 1 -Start ([Datetime]"08:00")

New-CMSoftwareUpdateAutoDeploymentRule `
    -CollectionName $Collection `
    -DeploymentPackageName $ADRName `
    -Name $ADRName `
    -AddToExistingSoftwareUpdateGroup $True `
    -AllowRestart $False `
    -AllowSoftwareInstallationOutsideMaintenanceWindow $False `
    -AllowUseMeteredNetwork $False `
    -AvailableImmediately $True `
    -DeadlineImmediately $True `
    -DeployWithoutLicense $False `
    -DisableOperationManager $True `
    -DownloadFromInternet $True `
    -DownloadFromMicrosoftUpdate $False `
    -EnabledAfterCreate $True `
    -GenerateOperationManagerAlert $True `
    -Language "English" `
    -LanguageSelection "English" `
    -NoInstallOnRemote $False `
    -NoInstallOnUnprotected $True `
    -Product $Products `
    -Required ">0" `
    -RunType RunTheRuleOnSchedule `
    -Schedule $Schedule `
    -SendWakeUpPacket $True `
    -Severity $Severity `
    -Superseded $False `
    -SuppressRestartServer $False `
    -SuppressRestartWorkstation $False `
    -Title "Preview" `
    -UpdateClassification $UpdateClassifications `
    -UserNotification DisplayAll `
    -UseUtc $False `
    -VerboseLevel AllMessages `
    -WriteFilterHandling $True `
}
 