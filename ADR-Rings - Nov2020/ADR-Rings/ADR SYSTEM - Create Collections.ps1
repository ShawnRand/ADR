#############################################################################
#
#
#############################################################################

#Load ConfigMge PoSH Module
Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

#Branding (change to customer name)
$OrgName = 'MSFT Services'

#SiteCode
$SiteCode = Get-PSDrive -PSProvider CMSITE
Set-location $SiteCode":"

#Error Handling and output
Clear-Host
$ErrorActionPreference= 'Continue'
#$Error1 = 0

#Refresh Schedule 
$Schedule = New-CMSchedule –RecurInterval Days –RecurCount 1

#DFG (DogFood Group) Collections Query
$CollectionDfG = @{Name = "*DogFood ADR Ring"}

$CollectionDfG1 = @{Name = "DogFood ADR Office 2013 2016 2019 365"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation%'"}
$CollectionDfG2 = @{Name = "DogFood ADR Windows 7"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.1%'"}
$CollectionDfG3 = @{Name = "DogFood ADR Windows 8 and 8.1"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.2%' or OperatingSystemNameandVersion like '%Workstation 6.3%'"}
$CollectionDfG4 = @{Name = "DogFood ADR Windows 10"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 10.0%'"}
$CollectionDfG5 = @{Name = "DogFood ADR Windows 2008 and 2008 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.0%' or OperatingSystemNameandVersion like '%Server 6.1%'"}
$CollectionDfG6 = @{Name = "DogFood ADR Windows 2012 and 2012 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.2%' or OperatingSystemNameandVersion like '%Server 6.3%'"}
$CollectionDfG7 = @{Name = "DogFood ADR Windows 2016 and 2019"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 10%'"}


#PlG (Pilot Group) Collections Query
$CollectionPlG = @{Name = "*Pilot ADR Ring"}

$CollectionPlG1 = @{Name = "Pilot ADR Office 2013 2016 2019 365"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation%'"}
$CollectionPlG2 = @{Name = "Pilot ADR Windows 7"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.1%'"}
$CollectionPlG3 = @{Name = "Pilot ADR Windows 8 and 8.1"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.2%' or OperatingSystemNameandVersion like '%Workstation 6.3%'"}
$CollectionPlG4 = @{Name = "Pilot ADR Windows 10"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 10.0%'"}
$CollectionPlG5 = @{Name = "Pilot ADR Windows 2008 and 2008 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.0%' or OperatingSystemNameandVersion like '%Server 6.1%'"}
$CollectionPlG6 = @{Name = "Pilot ADR Windows 2012 and 2012 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.2%' or OperatingSystemNameandVersion like '%Server 6.3%'"}
$CollectionPlG7 = @{Name = "Pilot ADR Windows 2016 and 2019"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 10%'"}

#PrG (Production Group) Collections Query
$CollectionPrG = @{Name = "*Production ADR Ring"}

$CollectionPrG1 = @{Name = "Production ADR Office 2013 2016 2019 365"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation%'"}
$CollectionPrG2 = @{Name = "Production ADR Windows 7"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.1%'"}
$CollectionPrG3 = @{Name = "Production ADR Windows 8 and 8.1"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 6.2%' or OperatingSystemNameandVersion like '%Workstation 6.3%'"}
$CollectionPrG4 = @{Name = "Production ADR Windows 10"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Workstation 10.0%'"}
$CollectionPrG5 = @{Name = "Production ADR Windows 2008 and 2008 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.0%' or OperatingSystemNameandVersion like '%Server 6.1%'"}
$CollectionPrG6 = @{Name = "Production ADR Windows 2012 and 2012 R2"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 6.2%' or OperatingSystemNameandVersion like '%Server 6.3%'"}
$CollectionPrG7 = @{Name = "Production ADR Windows 2016 and 2019"; Query = "select SMS_R_System.ResourceID,SMS_R_System.ResourceType,SMS_R_System.Name,SMS_R_System.SMSUniqueIdentifier,SMS_R_System.ResourceDomainORWorkgroup,SMS_R_System.Client from SMS_R_System where OperatingSystemNameandVersion like '%Server 10%'"}

#Create Root Folder
$CollectionFolder = @{Name = "$OrgName ADR Collections"; ObjectType = 5000; ParentContainerNodeId = 0}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $CollectionFolder

#Find Root Folder
#$ParentFolderID  = Get-wmiObject -Namespace root\SMS\site_$SCCMSiteCode -Query "Select containernodeid from SMS_ObjectContainerNode" | select ContainerNodeID | Where-Object {$_.Name -eq $CollectionFolder}
$ParentFolderID  = Get-wmiObject -Namespace root\SMS\site_$SiteCode -Query "Select * from SMS_ObjectContainerNode Where Name = '$OrgName ADR Collections'"

#write-host $ParentFolderID.ContainerNodeID
#write-host $ParentFolderID.name

#Create Sub Folders
$DogFoodCollectionFolder = @{Name = "$OrgName DogFood ADR Collections"; ObjectType = 5000; ParentContainerNodeId = $ParentFolderID.ContainerNodeID}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $DogFoodCollectionFolder

$PilotCollectionFolder = @{Name = "$OrgName Pilot ADR Collections"; ObjectType = 5000; ParentContainerNodeId = $ParentFolderID.ContainerNodeID}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $PilotCollectionFolder

$ProdCollectionFolder = @{Name = "$OrgName Prod ADR Collections"; ObjectType = 5000; ParentContainerNodeId = $ParentFolderID.ContainerNodeID}
Set-WmiInstance -Namespace "root\sms\site_$($SiteCode.Name)" -Class "SMS_ObjectContainerNode" -Arguments $ProdCollectionFolder

#Limiting collections
$AllSystems = "All Systems"
$DogFoodLimitingCollection = "$OrgName ADR Dogfood"
$PilotLimitingCollection = "$OrgName ADR Pilot"
$ProdLimitingCollection = "$OrgName ADR Prod"

#Create Collection


#Base Collections
New-CMDeviceCollection -Name $CollectionDFG.Name -Comment "DogFood System" -LimitingCollectionName $AllSystems -RefreshSchedule $Schedule -RefreshType Both | Out-Null
Write-host *** Collection $CollectionDFG.Name created ***

New-CMDeviceCollection -Name $CollectionPlG.Name -Comment "Pilot System" -LimitingCollectionName $AllSystems -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Write-host *** Collection $CollectionPlG.Name created ***

New-CMDeviceCollection -Name $CollectionPrG.Name -Comment "Prodution System" -LimitingCollectionName $AllSystems -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Write-host *** Collection $CollectionPrG.Name created ***

#DogFood Collections
New-CMDeviceCollection -Name $CollectionDfG1.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG1.Name -QueryExpression $CollectionDfG1.Query -RuleName $CollectionDfG1.Name
Write-host *** Collection $CollectionDfG1.Name created ***

New-CMDeviceCollection -Name $CollectionDfG2.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG2.Name -QueryExpression $CollectionDfG2.Query -RuleName $CollectionDfG2.Name
Write-host *** Collection $CollectionDfG2.Name created ***

New-CMDeviceCollection -Name $CollectionDfG3.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG3.Name -QueryExpression $CollectionDfG3.Query -RuleName $CollectionDfG3.Name
Write-host *** Collection $CollectionDfG3.Name created ***

New-CMDeviceCollection -Name $CollectionDfG4.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG4.Name -QueryExpression $CollectionDfG4.Query -RuleName $CollectionDfG4.Name
Write-host *** Collection $CollectionDfG4.Name created ***

New-CMDeviceCollection -Name $CollectionDfG5.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG5.Name -QueryExpression $CollectionDfG5.Query -RuleName $CollectionDfG5.Name
Write-host *** Collection $CollectionDfG5.Name created ***

New-CMDeviceCollection -Name $CollectionDfG6.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG6.Name -QueryExpression $CollectionDfG6.Query -RuleName $CollectionDfG6.Name
Write-host *** Collection $CollectionDfG6.Name created ***

New-CMDeviceCollection -Name $CollectionDfG7.Name -Comment "" -LimitingCollectionName $CollectionDFG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionDfG7.Name -QueryExpression $CollectionDfG7.Query -RuleName $CollectionDfG7.Name
Write-host *** Collection $CollectionDfG7.Name created ***

#Pilot Collections

New-CMDeviceCollection -Name $CollectionPlG1.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG1.Name -QueryExpression $CollectionPlG1.Query -RuleName $CollectionPlG1.Name
Write-host *** Collection $CollectionPlG1.Name created ***

New-CMDeviceCollection -Name $CollectionPlG2.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG2.Name -QueryExpression $CollectionPlG2.Query -RuleName $CollectionPlG2.Name
Write-host *** Collection $CollectionPlG2.Name created ***

New-CMDeviceCollection -Name $CollectionPlG3.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG3.Name -QueryExpression $CollectionPlG3.Query -RuleName $CollectionPlG3.Name
Write-host *** Collection $CollectionPlG3.Name created ***

New-CMDeviceCollection -Name $CollectionPlG4.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG4.Name -QueryExpression $CollectionPlG4.Query -RuleName $CollectionPlG4.Name
Write-host *** Collection $CollectionPlG4.Name created ***

New-CMDeviceCollection -Name $CollectionPlG5.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG5.Name -QueryExpression $CollectionPlG5.Query -RuleName $CollectionPlG5.Name
Write-host *** Collection $CollectionPlG5.Name created ***

New-CMDeviceCollection -Name $CollectionPlG6.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG6.Name -QueryExpression $CollectionPlG6.Query -RuleName $CollectionPlG6.Name
Write-host *** Collection $CollectionPlG6.Name created ***

New-CMDeviceCollection -Name $CollectionPlG7.Name -Comment "" -LimitingCollectionName $CollectionPlG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPlG7.Name -QueryExpression $CollectionPlG7.Query -RuleName $CollectionPlG7.Name
Write-host *** Collection $CollectionPlG7.Name created ***

#Prod Collections

New-CMDeviceCollection -Name $CollectionPrG1.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG1.Name -QueryExpression $CollectionPrG1.Query -RuleName $CollectionPrG1.Name
Write-host *** Collection $CollectionPrG1.Name created ***

New-CMDeviceCollection -Name $CollectionPrG2.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG2.Name -QueryExpression $CollectionPrG2.Query -RuleName $CollectionPrG2.Name
Write-host *** Collection $CollectionPrG2.Name created ***

New-CMDeviceCollection -Name $CollectionPrG3.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG3.Name -QueryExpression $CollectionPrG3.Query -RuleName $CollectionPrG3.Name
Write-host *** Collection $CollectionPrG3.Name created ***

New-CMDeviceCollection -Name $CollectionPrG4.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG4.Name -QueryExpression $CollectionPrG4.Query -RuleName $CollectionPrG4.Name
Write-host *** Collection $CollectionPrG4.Name created ***

New-CMDeviceCollection -Name $CollectionPrG5.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG5.Name -QueryExpression $CollectionPrG5.Query -RuleName $CollectionPrG5.Name
Write-host *** Collection $CollectionPrG5.Name created ***

New-CMDeviceCollection -Name $CollectionPrG6.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG6.Name -QueryExpression $CollectionPrG6.Query -RuleName $CollectionPrG6.Name
Write-host *** Collection $CollectionPrG6.Name created ***

New-CMDeviceCollection -Name $CollectionPrG7.Name -Comment "" -LimitingCollectionName $CollectionPrG.Name -RefreshSchedule $Schedule -RefreshType Both  | Out-Null
Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionPrG7.Name -QueryExpression $CollectionPrG7.Query -RuleName $CollectionPrG7.Name
Write-host *** Collection $CollectionPrG7.Name created ***

#Move the DogFood collections to the right folder
$FolderPath = $SiteCode.Name + ":\DeviceCollection\" + $CollectionFolder.Name +"\"+ $DogFoodCollectionFolder.name
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG1.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG2.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG3.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG4.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG5.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG6.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionDFG7.Name)


#Move the Pilot collections to the right folder
$FolderPath = $SiteCode.Name + ":\DeviceCollection\" + $CollectionFolder.Name +"\"+ $PilotCollectionFolder.name
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG1.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG2.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG3.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG4.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG5.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG6.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPlG7.Name)


#Move the Prod collections to the right folder
$FolderPath = $SiteCode.Name + ":\DeviceCollection\" + $CollectionFolder.Name +"\"+ $ProdCollectionFolder.name
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG1.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG2.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG3.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG4.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG5.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG6.Name)
Move-CMObject -FolderPath $FolderPath -InputObject (Get-CMDeviceCollection -Name $CollectionPrG7.Name)
