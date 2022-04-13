#############################################################################
#
#
#############################################################################

#Load ConfigMge PoSH Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Branding (change to customer name)
$OrgName = "MSFT Services"

#SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference= 'Continue'
#$Error1 = 0

#Refresh Schedule 
$Schedule = New-CMSchedule –RecurInterval Days –RecurCount 1

#PvG (Preview Group) Collections Query
$CollectionPvG = @{Name = "*Preview ADR Ring"}

$CollectionPvG1 = @{Name = "Preview ADR Office 2013 2016 2019 365"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation%'"}
$CollectionPvG2 = @{Name = "Preview ADR Windows 7"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.1%'"}
$CollectionPvG3 = @{Name = "Preview ADR Windows 8 and 8.1"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.2%' or OperatingSystemNameandVersion like '%Workstation 6.3%'"}
$CollectionPvG4 = @{Name = "Preview ADR Windows 10"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 10.0%'"}
$CollectionPvG5 = @{Name = "Preview ADR Windows 2008 and 2008 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.0%' or OperatingSystemNameandVersion like '%Server 6.1%'"}
$CollectionPvG6 = @{Name = "Preview ADR Windows 2012 and 2012 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.2%' or OperatingSystemNameandVersion like '%Server 6.3%'"}
$CollectionPvG7 = @{Name = "Preview ADR Windows 2016 and 2019"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 10%'"}

#Create Root Folder
$CollectionFolder = @{Name = "$OrgName ADR Collections"; ObjectType = 5000; ParentContainerNodeId = 0}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolder

#Find Root Folder
#$ParentFolderID  = Get-wmiObject -Namespace root\SMS\site_$SCCMSiteCode -Query "Select containernodeid from SMS_ObjectContainerNode" | select ContainerNodeID | Where-Object {$_.Name -eq $CollectionFolder}
$ParentFolderID  = Get-wmiObject -Namespace root\SMS\site_$SiteCode -Query "Select * from SMS_ObjectContainerNode Where Name = '$OrgName ADR Collections'"

#write-host $ParentFolderID.ContainerNodeID
#write-host $ParentFolderID.name

#Create Sub Folders
$PreviewCollectionFolder = @{Name = "$OrgName Preview ADR Collections"; ObjectType = 5000; ParentContainerNodeId = $ParentFolderID.ContainerNodeID}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $PreviewCollectionFolder

#Limiting collections
$AllSystems = "All Systems"
$PreviewLimitingCollection = "$OrgName ADR Preview"

#Create Collection

#Base Collections
New-CMDeviceCollection -Name $CollectionPvG.Name -Comment "Preview System" -LimitingCollectionName $AllSystems -RefreshSchedule $Schedule -RefreshType Both | Out-Null
Write-host *** Collection $CollectionPvG.Name created ***


#Preview Collections
New-CMDeviceCollection -Name $CollectionPvG1.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPVG1.Name -QueryExpression $CollectionPvG1.Query -RuleName $CollectionPvG1.Name
Write-host *** Collection $CollectionPvG1.Name created ***

New-CMDeviceCollection -Name $CollectionPvG2.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG2.Name -QueryExpression $CollectionPvG2.Query -RuleName $CollectionPvG2.Name
Write-host *** Collection $CollectionPvG2.Name created ***

New-CMDeviceCollection -Name $CollectionPvG3.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG3.Name -QueryExpression $CollectionPvG3.Query -RuleName $CollectionPvG3.Name
Write-host *** Collection $CollectionPvG3.Name created ***

New-CMDeviceCollection -Name $CollectionPvG4.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG4.Name -QueryExpression $CollectionPvG4.Query -RuleName $CollectionPvG4.Name
Write-host *** Collection $CollectionPvG4.Name created ***

New-CMDeviceCollection -Name $CollectionPvG5.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG5.Name -QueryExpression $CollectionPvG5.Query -RuleName $CollectionPvG5.Name
Write-host *** Collection $CollectionPvG5.Name created ***

New-CMDeviceCollection -Name $CollectionPvG6.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG6.Name -QueryExpression $CollectionPvG6.Query -RuleName $CollectionPvG6.Name
Write-host *** Collection $CollectionPvG6.Name created ***

New-CMDeviceCollection -Name $CollectionPvG7.Name -Comment "" -LimitingCollectionName $CollectionPvG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPvG7.Name -QueryExpression $CollectionPvG7.Query -RuleName $CollectionPvG7.Name
Write-host *** Collection $CollectionPvG7.Name created ***


#Move the Preview collections to the right folder
$FolderPath = $SiteCode.Name + ":\DeviceCollection\" + $CollectionFolder.Name +"\"+ $PreviewCollectionFolder.name
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG1.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG2.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG3.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG4.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG5.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG6.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPvG7.Name)
