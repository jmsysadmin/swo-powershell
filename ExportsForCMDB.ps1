<#	
	.NOTES
	===========================================================================
	 Created on:   	11/9/2021 9:47 AM
	 Created by:   	Jake Muszynski
	 Organization: 	Nationwide Children's Hospital
	 Filename:     	ExportsForCMDB.ps1
	===========================================================================
	.DESCRIPTION
		This script connects to Solarwinds Orion to pull out information that 
		can be imported into a CMDB in csv files.
#>

Function Get-ScriptDirectory {
	[OutputType([string])]
	Param ()
	If ($null -ne $hostinvocation) {
		Split-Path $hostinvocation.MyCommand.path
	} Else {
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

Try {
	Import-Module SwisPowerShell
} Catch {
	Write-Host "Message: $($Error[0])";
	Exit 1;
}

[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$SwisServer = [Microsoft.VisualBasic.Interaction]::InputBox("What Orion Server Are you connecting to?", "Orion Server Name")

Try {
	$swis = Connect-Swis $SwisServer -Trusted 
} Catch {
	$SwisServer = [Microsoft.VisualBasic.Interaction]::InputBox("What Orion Server Are you connecting to?", "Orion Server Name")
	$Creds = Get-Credential
	$Swis = Connect-Swis $SwisServer -Credential $Creds
}


$scriptfolder = Get-ScriptDirectory

$Server_AIX_query = "SELECT 
    Case 
        When N.AssetInventory.ServerInformation.HostName is not Null Then N.AssetInventory.ServerInformation.HostName
        When N.NodeName is not Null Then N.NodeName
        Else N.Caption 
    End as [host_name]
,   Case 
    When N.AssetInventory.ServerInformation.Domain is not Null AND N.AssetInventory.ServerInformation.Domain not like '' and N.AssetInventory.ServerInformation.HostName is not Null then (N.AssetInventory.ServerInformation.HostName + '.' + N.AssetInventory.ServerInformation.Domain)
    When N.DNS is not Null Then N.DNS
    End as [fqdn]
,   Case 
        When N.AssetInventory.ServerInformation.Manufacturer is Null Then 'IBM'
        Else N.AssetInventory.ServerInformation.Manufacturer
    End as [manufacturer display_value]
,   Case 
        When N.AssetInventory.ServerInformation.OperatingSystem like '' Then 'AIX'
        Else N.AssetInventory.ServerInformation.OperatingSystem
    End as [os]
,   N.AssetInventory.ServerInformation.OSVersion as [os_version]
,   N.AssetInventory.ServerInformation.OSArchitecture as [os_address_width] 
,   N.AssetInventory.ServerInformation.ServicePack as [os_service_pack]
,   N.AssetInventory.ServerInformation.Domain as [os_domain]
,   N.AssetInventory.ServerInformation.HardwareSerialNumber as [serial_number]
,   N.AssetInventory.ServerInformation.TotalMemoryB as [Ram]
,   Case 
        When N.AssetInventory.ServerInformation.Manufacturer like 'VMware, Inc.' Then 'True'
        When N.AssetInventory.ServerInformation.Manufacturer like 'Microsoft Corporation' Then 'True'
        Else 'False'
    End as [virtual]
,   N.NodeID 
FROM Orion.Nodes N 
where N.MachineType like 'IBM%'"
$Server_Esxi_query = "SELECT 
    Case 
        When N.AssetInventory.ServerInformation.HostName is not Null Then N.AssetInventory.ServerInformation.HostName
        When N.NodeName is not Null Then N.NodeName
        Else N.Caption 
    End as [host_name]
,   n.Description
,   Case 
    When N.AssetInventory.ServerInformation.Domain is not Null AND N.AssetInventory.ServerInformation.Domain not like '' and N.AssetInventory.ServerInformation.HostName is not Null then (N.AssetInventory.ServerInformation.HostName + '.' + N.AssetInventory.ServerInformation.Domain)
    When N.DNS is not Null Then N.DNS
    End as [fqdn]
,   N.AssetInventory.ServerInformation.Manufacturer as [manufacturer display_value]
,   N.AssetInventory.ServerInformation.OperatingSystem as [os]
,   N.AssetInventory.ServerInformation.OSVersion as [os_version]
,   N.AssetInventory.ServerInformation.OSArchitecture as [os_address_width] 
,   N.AssetInventory.ServerInformation.ServicePack as [os_service_pack]
,   N.AssetInventory.ServerInformation.Domain as [os_domain]
,   N.AssetInventory.ServerInformation.HardwareSerialNumber as [serial_number]
,   N.AssetInventory.ServerInformation.TotalMemoryB as [Ram]
,   Case 
        When N.AssetInventory.ServerInformation.Manufacturer like 'VMware, Inc.' Then 'True'
        When N.AssetInventory.ServerInformation.Manufacturer like 'Microsoft Corporation' Then 'True'
        Else 'False'
    End as [virtual]
,   N.NodeID 
FROM Orion.Nodes N 
where N.AssetInventory.ServerInformation.OperatingSystem like 'VMware ESXi'"
$Server_Linux_query = "SELECT 
    Case 
        When N.AssetInventory.ServerInformation.HostName is not Null Then N.AssetInventory.ServerInformation.HostName
        When N.NodeName is not Null Then N.NodeName
        Else N.Caption 
    End as [host_name]
,   n.Description	
,   Case 
    When N.AssetInventory.ServerInformation.Domain is not Null AND N.AssetInventory.ServerInformation.Domain not like '' and N.AssetInventory.ServerInformation.HostName is not Null then (N.AssetInventory.ServerInformation.HostName + '.' + N.AssetInventory.ServerInformation.Domain)
    When N.DNS is not Null Then N.DNS
    End as [fqdn]
,   N.AssetInventory.ServerInformation.Manufacturer as [manufacturer display_value]
,   N.AssetInventory.ServerInformation.OperatingSystem as [os]
,   N.AssetInventory.ServerInformation.OSVersion as [os_version]
,   N.AssetInventory.ServerInformation.OSArchitecture as [os_address_width] 
,   N.AssetInventory.ServerInformation.ServicePack as [os_service_pack]
,   N.AssetInventory.ServerInformation.Domain as [os_domain]
,   N.AssetInventory.ServerInformation.HardwareSerialNumber as [serial_number]
,   N.AssetInventory.ServerInformation.TotalMemoryB as [Ram]
,   Case 
        When N.AssetInventory.ServerInformation.Manufacturer like 'VMware, Inc.' Then 'True'
        When N.AssetInventory.ServerInformation.Manufacturer like 'Microsoft Corporation' Then 'True'
        Else 'False'
    End as [virtual]
,   N.NodeID 
FROM Orion.Nodes N 
where 
	N.IsServer = True
	and N.MachineType like 'net-snmp - Linux'"
$Server_Windows_query = "SELECT 
    Case 
        When N.AssetInventory.ServerInformation.HostName is not Null Then N.AssetInventory.ServerInformation.HostName
        When N.NodeName is not Null Then N.NodeName
        Else N.Caption 
    End as [host_name]
,   Case 
    When N.AssetInventory.ServerInformation.Domain is not Null AND N.AssetInventory.ServerInformation.Domain not like '' and N.AssetInventory.ServerInformation.HostName is not Null then (N.AssetInventory.ServerInformation.HostName + '.' + N.AssetInventory.ServerInformation.Domain)
    When N.DNS is not Null Then N.DNS
    End as [fqdn]
,   N.AssetInventory.ServerInformation.Manufacturer as [manufacturer display_value]
,   N.AssetInventory.ServerInformation.OperatingSystem as [os]
,   N.AssetInventory.ServerInformation.OSVersion as [os_version]
,   N.AssetInventory.ServerInformation.OSArchitecture as [os_address_width] 
,   N.AssetInventory.ServerInformation.ServicePack as [os_service_pack]
,   N.AssetInventory.ServerInformation.Domain as [os_domain]
,   N.AssetInventory.ServerInformation.HardwareSerialNumber as [serial_number]
,   N.AssetInventory.ServerInformation.TotalMemoryB as [Ram]
,   Case 
        When N.AssetInventory.ServerInformation.Manufacturer like 'VMware, Inc.' Then 'True'
        When N.AssetInventory.ServerInformation.Manufacturer like 'Microsoft Corporation' Then 'True'
        Else 'False'
    End as [virtual]
,   N.NodeID 
FROM Orion.Nodes N 
where N.MachineType like 'Windows%Server%' or N.MachineType like 'Windows%Domain Controller%'"
$Network_Device_query = "SELECT
  NCM_EP.Node.NodeCaption
, NCM_EP.Node.AgentIP
, NCM_EP.Node.MachineType
, Case 
    When NCM_EP.Manufacturer like '' and NCM_EP.Node.MachineType like '%Cisco%'  Then 'Cisco Systems Inc'
    When NCM_EP.Manufacturer like '' and NCM_EP.Node.MachineType like 'Catalyst%'  Then 'Cisco Systems Inc'
    When NCM_EP.Manufacturer is Null and NCM_EP.Node.MachineType like 'Catalyst%'  Then 'Cisco Systems Inc'
    When NCM_EP.Manufacturer like 'Cisco' Then 'Cisco Systems Inc'
    Else NCM_EP.Manufacturer
End as [ManufacturerCalc]
, Model
, EntityName
, EntityDescription
, HardwareRevision
, Serial

FROM NCM.EntityPhysical NCM_EP 
Where 
    Serial not like '' 
    and EntityClass not like '1' 
    and EntityClass not like '10' 
    and EntityClass not like '6' 
    and EntityClass not like '9' 
    and EntityClass not like '5'
    and EntityClass not like '7'"
$Vmware_Hosts_query = "SELECT 
    N.VCenter.DataCenters.Clusters.Hosts.HostID
,   N.VCenter.DataCenters.Clusters.Hosts.NodeID
,   N.VCenter.DataCenters.Clusters.Hosts.HostName
,   N.VCenter.DataCenters.Clusters.Hosts.ClusterID
,   N.VCenter.DataCenters.Clusters.Hosts.SecondaryClusterID
,   N.VCenter.DataCenters.Clusters.Hosts.DataCenterID
,   N.VCenter.DataCenters.Clusters.Hosts.VMwareProductName
,   N.VCenter.DataCenters.Clusters.Hosts.VMwareProductVersion
,   N.VCenter.DataCenters.Clusters.Hosts.Model
,   N.VCenter.DataCenters.Clusters.Hosts.Vendor
,   N.VCenter.DataCenters.Clusters.Hosts.ProcessorType
,   N.VCenter.DataCenters.Clusters.Hosts.CpuCoreCount
,   N.VCenter.DataCenters.Clusters.Hosts.CpuPkgCount
,   N.VCenter.DataCenters.Clusters.Hosts.CpuMhz
,   N.VCenter.DataCenters.Clusters.Hosts.NicCount
,   N.VCenter.DataCenters.Clusters.Hosts.HbaCount
,   N.VCenter.DataCenters.Clusters.Hosts.HbaFcCount
,   N.VCenter.DataCenters.Clusters.Hosts.HbaScsiCount
,   N.VCenter.DataCenters.Clusters.Hosts.HbaIscsiCount
,   N.VCenter.DataCenters.Clusters.Hosts.MemorySize
,   N.VCenter.DataCenters.Clusters.Hosts.BiosSerial
,   N.VCenter.DataCenters.Clusters.Hosts.DNSHostName
,   N.VCenter.DataCenters.Clusters.Hosts.IPAddress
,   N.VCenter.DataCenters.Clusters.Hosts.VmCount
,   N.VCenter.DataCenters.Clusters.Hosts.VmRunningCount
FROM Orion.Nodes N 
Where N.VCenter.DataCenters.Clusters.Hosts.HostID > 0"
$Vmware_Clusters_query = "SELECT 
    N.VCenter.DataCenters.Clusters.ClusterID
,   N.VCenter.DataCenters.Clusters.DataCenterID
,   N.VCenter.DataCenters.Clusters.Name
,   N.VCenter.DataCenters.Clusters.TotalMemory
,   N.VCenter.DataCenters.Clusters.TotalCpu
,   N.VCenter.DataCenters.Clusters.CpuCoreCount
,   N.VCenter.DataCenters.Clusters.CpuThreadCount
,   N.VCenter.DataCenters.Clusters.EffectiveCpu
,   N.VCenter.DataCenters.Clusters.EffectiveMemory
,   N.VCenter.DataCenters.Clusters.DatastoreUsedSpace
FROM Orion.Nodes N 
Where N.VCenter.DataCenters.Clusters.ClusterID > 0"
$Vmware_DataCenters_query = "SELECT 
    N.VCenter.DataCenters.DataCenterID
,   N.VCenter.DataCenters.ManagedObjectID
,   N.VCenter.DataCenters.VCenterID
,   N.VCenter.DataCenters.Name
,   N.VCenter.DataCenters.ConfigStatus
,   N.VCenter.DataCenters.OverallStatus
,   N.VCenter.DataCenters.ManagedStatus
,   N.VCenter.DataCenters.TriggeredAlarmDescription
,   N.VCenter.DataCenters.PollingSource
,   N.VCenter.DataCenters.OrionIdPrefix
,   N.VCenter.DataCenters.OrionIdColumn
,   N.VCenter.DataCenters.DetailsUrl
FROM Orion.Nodes N 
Where N.VCenter.DataCenters.DataCenterID > 0"
$Vmware_Datastores_query = "SELECT 
    N.VCenter.DataCenters.Clusters.DataStores.DataStoreID
,   N.VCenter.DataCenters.Clusters.DataStores.DataStoreIdentifier
,   N.VCenter.DataCenters.Clusters.DataStores.Name
,   N.VCenter.DataCenters.Clusters.DataStores.Type
,   N.VCenter.DataCenters.Clusters.DataStores.Capacity
,   N.VCenter.DataCenters.Clusters.DataStores.FreeSpace
,   N.VCenter.DataCenters.Clusters.DataStores.ProvisionedSpace
,   N.VCenter.DataCenters.Clusters.DataStores.SpaceUtilization
,   N.VCenter.DataCenters.Clusters.DataStores.ProvisionedSpaceAllocation
,   N.VCenter.VCenterID
,   N.VCenter.DataCenters.Clusters.ClusterID
FROM Orion.Nodes N 
Where N.VCenter.DataCenters.Clusters.DataStores.DataStoreID > 0"
$Vmware_relationships_query = "SELECT 
    N.VCenter.VCenterID
,   N.VCenter.DataCenters.DataCenterID
,   N.VCenter.DataCenters.Clusters.ClusterID
,    N.VCenter.DataCenters.Clusters.DataStores.DataStoreID
FROM Orion.Nodes N 
Where N.VCenter.DataCenters.Clusters.DataStores.DataStoreID > 0"

$Server_AIX_result           = Get-SwisData $swis $Server_AIX_query
$Server_Esxi_result          = Get-SwisData $swis $Server_Esxi_query
$Server_Linux_result         = Get-SwisData $swis $Server_Linux_query
$Server_Windows_result       = Get-SwisData $swis $Server_Windows_query
$Network_Device_result       = Get-SwisData $swis $Network_Device_query
$Vmware_Hosts_result         = Get-SwisData $swis $Vmware_Hosts_query
$Vmware_Clusters_result      = Get-SwisData $swis $Vmware_Clusters_query
$Vmware_DataCenters_result   = Get-SwisData $swis $Vmware_DataCenters_query
$Vmware_Datastores_result    = Get-SwisData $swis $Vmware_Datastores_query
$Vmware_relationships_result = Get-SwisData $swis $Vmware_relationships_query

$Server_AIX_result | Export-Csv -Path "$scriptfolder\Server_AIX.csv"
$Server_Esxi_result | Export-Csv -Path "$scriptfolder\Server_Esxi.csv"
$Server_Linux_result | Export-Csv -Path "$scriptfolder\Server_Linux.csv"
$Server_Windows_result | Export-Csv -Path "$scriptfolder\Server_Windows.csv"
$Network_Device_result | Export-Csv -Path "$scriptfolder\Network_Devices.csv"
$Vmware_Hosts_result | Export-Csv -Path "$scriptfolder\Vmware_Hosts.csv"
$Vmware_Clusters_result | Export-Csv -Path "$scriptfolder\Vmware_Clusters.csv"
$Vmware_DataCenters_result | Export-Csv -Path "$scriptfolder\Vmware_DataCenters.csv"
$Vmware_Datastores_result | Export-Csv -Path "$scriptfolder\Vmware_Datastores.csv"
$Vmware_relationships_result | Export-Csv -Path "$scriptfolder\Vmware_relationships.csv"
