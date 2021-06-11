#################################################
#
# Function : Check-ForImportExcel
#
# Author : Dave Lloyd
#
# Purpose : Check if the ImportExcel module is
# available for use. If so, we can create .xlsx
# files, otherwise we either have to try and 
# import it, or fall back to having to generate
# .csv files only 
#
#################################################
Function Test-ForImportExcel {
    $CurrentlyAvailableModules = (Get-Module -ListAvailable | Where-Object { $_.Name -eq "ImportExcel" })
    If ($CurrentlyAvailableModules) {
        Write-Host "ImportExcel module is available" -ForegroundColor Green
        Return $True
    } else {
        Return $False
    }
} # End Check-ForImportExcel

#################################################
#
# Function : Set-InitializeAudit
#
# Author : Dave Lloyd
#
# Purpose : Performs "initialization" when the
# script is run - check to see if the folder
# structure exists for reporting - if not, create it.
#
#################################################
Function Set-InitializeAudit {
    $temp_home = [Environment]::GetFolderPath("MyDocuments") # Get the MyDocuments folder - we'll be creating a folder here for output. Using environment supports roaming profile path.
    $Global:EnvironmentReport_home = $temp_home + "\EnvironmentReports" # This is going to be our default for logging output. Remainder will be subfolders to this.
    $Global:Reports_home = $EnvironmentReport_home + "\Reports\" # What follows will be subfolders for different types of audits.
    
    # If the folder for logging output doesn't exist, create it and the subfolders.
    Clear-Host
    Write-Host "Initialization checks in progress." -ForegroundColor Green
    Write-Host "----------------------------------" -ForegroundColor Green
    Write-Host "`nChecking if input and report folders exist." -ForegroundColor Green
    If (!(Test-Path -Path $Global:EnvironmentReport_home)) {
        Write-Host "Folders not present. Creating." -ForegroundColor Green
        New-Item -ItemType directory -Path $Global:EnvironmentReport_home
    } else {
        Write-Host "Folder present." -ForegroundColor Green
    }
    
    Write-Host "`nCheck if ImportExcel module is available." -ForegroundColor Green
    Test-ForImportExcel
        
    Write-Host "`nInitialization checks complete." -ForegroundColor Green
    Read-Host "`nPress ENTER to continue."
} # End Set-InitializeAudit
        
#################################################
#
# Function : Get-EnvironmentReport
#
# Author : Dave Lloyd
#
# Purpose : Main function for producing the audits.
#
# Version 0.1 April 2020 - initial release
#
#################################################
function Get-EnvironmentReport {    
    <#
        .Synopsis
         Generate an audit of a specific environment. This is a fairly basic audit.
        .DESCRIPTION
         Generates an audit of a specific generating output for the following properties based
         on the parameter value you provide - Yes or No

         The amount of information reported is determined by the ReportType parameter. The "Detailed" option generates the full report, and the "Summary" option is a subset of this. This applies to the properties reported in the VMs worksheet. 
         You can of course choose "Detailed" even for a summary report and just remove a few columns if they are not required.

         VCs
            Name
            Version
            Build
            
         Clusters 
            Cluster
            Overall status
            Configuration status
            HA Enabled  
            HA Admission control
            HA restart priority 
            HA isolation response
            DRS
            DRS mode
            DRS automation level
            Number vMotions
            vSAN enabled
            Number hosts
            CPU Cores
            CPU threads
            VM Count
            Powered off VM count

         VMs - Information included for "Summary" report type :
            VM 
            Powerstate
            CPUs
            Memory 
            Provisioned space
            Guest OS full name 
            Guest Name - guestname at the OS level - requires that VMwareTools be running.

         VMs - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
            Individual Hard Disk sizes (GB) - the size of the harddisks presented to the VM.
            Volume/Partition : Capacity : Free - disk usage at the OS level. Requires VMwareTools be running. All volumes/partitions in a single cell.
            IP Address(es) - requires VMwareTools be running. 
            MAC Addresses 
            NIC Connection State
            NIC Type - ie, vmxnet3 etc
            Portgroup name 
            Guest OS ID - VMware ID for the OS type   
            VMware Tools status
            VMware Tools version
            CPU hot add - True or False
            Memory hot add - True or False
            VM hardware version
            Host VM is running on
            Host version number
            Host build number

        VM Performance - only produced if "Detailed" report type is selected AND "Performance" equals Yes. Treat with caution.
            MemoryReservation - whether VM has a memory reservation and if so, what size. No reservation = 0
            MemoryLimit - whether VM has a memory limit and if so, what size. -1 = no limit.
            CPUReservation - whether VM has a CPU reservation and if so, what size. No reservation = 0
            CPULimit - whether VM has a CPU limit and if so, what size. -1 = no limit.
            Ballooning - value for the Summary.QuickStats.BalloonedMemory property
            Avg CPU Usage (Mhz) - derived from cpu.usagemhz.average metric over 30 days with 5 minute interval
            Avg Memory Usage (%) - derived from memory.usage.average metric over 30 days with 5 minute interval
            Avg Network Usage (KBps) - derived from network.usage.average metric over 30 days with 5 minute interval
            Avg Disk Usage (KBps) - derived from disk.usage.average metric over 30 days with 5 minute interval

            This information should probably be treated with caution, due to the sample frequency, and it simply being a single value representing the average over 30 days. Detailed performance information really needs something more thorough, such as vROPS.

        VM Disks - only produced if "Detailed" report type is selected :
            VM 
            HardDisk 
            Datastore 
            Size in GB

         RDMs - only produced if "Detailed" report type is selected :
            VM Name
            Disk Name
            Disk Type
            NAA
            VML         
            Filename   
            Capacity    

         DRS Rules - only produced if "Detailed" report type is selected :
            Cluster
            Rule Name
            Enabled - TRUE or FALSE
            Rule type - affinitiy/anti-affinity/VM host affinity
            VMs

         ESXi hosts - Information included for "Summary" report type :
            Host  
            Connection state 
            Boot time
            Host uptime in days 
            ESXi version number 
            Number sockets 
            Number of cores 
            Total CPU threads 
            Memory  
            Number of VMs running on the host 
            30 day maximum | minimum | average for CPU and Memory 

         ESXi hosts - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
            Vendor 
            Model  
            ESXi build number 
            CPU Model 

         ESXi NICs - Information included only if hosts AND "detailed" report type are selected :
            Host 
            NIC Name
            MAC Address
            Description
            Link status
            Link Speed
            Driver Type
            MTU

         ESXi vmks - only produced if hosts AND "detailed" report type are selected.
            Host
            vmk Name
            IP
            Subnet mask
            MAC
            Portgroup
            MTU
            Management - whether enabled for management purposes - TRUE or FALSE
            vMotion - whether enabled for vMotion - TRUE or FALSE
            FT - whether enabled for Fault Tolerance logging - TRUE or FALSE
            VSAN - whether enabled for VSAN - TRUE or FALSE

         vss Portgroups - only produced if hosts and "detailed" report type are selected.
            Host
            Portgroup Name
            Virtual Switch 
            VLAN ID

         vds Portgroups - only produced if hosts and "detailed" report type are selected.
            Host
            Portgroup Name
            Virtual Switch 
            Uplink
            VLAN ID
            PVLAN ID

         CDP - only produced if hosts and "detailed" report type are selected.
            Hostname
            VMNIC
            Speed
            PCI
            Port
            NetworkSwitch
            CDP_Address
            VLAN
            SwitchType

         Datastores - Information included for "Summary" report type :
            Datastore name 
            Capacity 
            Free space 
            Percentage free - values are rounded up. 

         Datastores - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
            Datastore name 
            State
            SIOC - whether enabled or not
            VMFS version
            NAA 
            ProvisionedGB - view to see if overprovisioned due to thin provisioning.

         Snapshots - Information included :
            VM 
            Snapshot name 
            Description 
            Date created 
            Created By - will be marked as "unknown" if unable to find a matching event.
            Snapshot age 
            Snapshot size 
            Datastore 
            Datastore free space 
         
         You will have the option to generate a single .xlsx file, which will have a worksheet per category, 
         or individual .csv files for each of the selected categories. 
         
         Select Excel or CSV for the OutputFormat parameter. The generation of the .xlsx file relies on the use 
         of the ImportExcel Powershell module. This will be checked for when the script runs.
         
         IF the check for the Excel module fails, the script will fall back to creating .csv files instead.

         Reports will be generated in a folder in your documents folder called EnvironmentReports 
         
        .PARAMETER vc
         The IP/FQDN of the VC to audit.
        .PARAMETER VMs
         Whether or not to include VMs in the audit.
        .PARAMETER Clusters
         Whether or not to include clusters in the audit.
        .PARAMETER Hosts
         Whether or not to include ESXi hosts in the audit.
        .PARAMETER Datastores
         Whether or not to include Datastores in the audit.
        .PARAMETER Snapshots
         Whether or not to include Snapshots in the audit.
        .PARAMETER OutputFormat
         Whether to ouput as an Excel or .CSV format
         Excel requires that the ImportExcel module be present.
         If the ImportExcel module isn't available, fallback to producing individual .csv files.
        .PARAMETER ReportType
         Summary will generate a smaller report, including only the most relevant/requested properties
         Detailed will effectively generate the full report.
        .PARAMETER Performance
         Whether to include "performance" data and worksheet if VMs are selected. This will increase the time for the script to run not insignificantly.
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate an .xlsx file Named 10.10.10.10-Audit-<date>.xlsx which contains worksheets for Clusters, VMs and performance, ESXi hosts and Datastores
         Get-EnvironmentReport -vCenter 10.10.10.10 -Clusters Yes -VMs Yes -Hosts Yes -Datastores Yes -Snapshots No -OutputFormat Excel -ReportType Details -Performance Yes    
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate an .xlsx file Named 10.10.10.10-Audit-<date>.xlsx which contains worksheets for VMs with no performance metrics, ESXi hosts and Datastores
         Get-EnvironmentReport -vCenter 10.10.10.10 -Clusters No  -VMs Yes -Hosts Yes -Datastores Yes -Snapshots No -OutputFormat Excel -ReportType Details -Performance No    
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate a series of .csv files, one each containing an audit of VMs, Clusters, ESXi hosts, Datastores and Snapshots in the environment. 
         Get-EnvironmentReport -vCenter 10.10.10.10 -Clusters Yes -VMs Yes -Hosts Yes -Datastores Yes -Snapshots Yes -OutputFormat CSV -ReportType Summary -Performance No
              
        .NOTES
        Author          : Dave Lloyd
        Version         : 0.1
      #>
    
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, Position = 1)]
        [string]$vCenter,
        
        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$Clusters,

        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$VMs,
    
        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$Hosts,
     
        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$Datastores,
     
        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$Snapshots,
     
        [Parameter(Mandatory = $True)]
        [ValidateSet('Excel', 'CSV')] # these are the only valid options
        [string]$OutputFormat,  

        [Parameter(Mandatory = $True)]
        [ValidateSet('Summary', 'Detailed')] # these are the only valid options
        [string]$ReportType,  

        [Parameter(Mandatory = $True)]
        [ValidateSet('Yes', 'No')] # these are the only valid options
        [string]$Performance
        
    )
    
    Set-InitializeAudit

    # Check to see if ImportExcel is available. If it's not, we'll have to default to .csv
    # We probably could/should offer to install the module instead - at least scoped for the current user.
    If ($OutputFormat -eq "Excel") {
        $ExcelAvailable = Test-ForImportExcel
        If (-not $ExcelAvailable) {
            Write-Host "ImportExcel appears to not be available. Switching output format to .csv" -ForegroundColor Green
            $OutputFormat = "CSV"
        }
    }
        
    # Log directory where the report will be generated.
    $script_dir = $Global:Reports_home
    
    # TP = time period. How many days info is going to get retrieved.
    $TP = "-30" 
    
    # Collections for each of the elements we are collecting for - Datacenter, hosts, VMs, datastores and VCs. 
    # If more is needed, eg, VC, then add an appropriate collection definition here.
    # The data presented in these will end up as a separate worksheet in the final .xlsx file
    $VMCollection = @() # VMs worksheet collection
    $VMPerfCollection = @() # VM Performance worksheet collection
    $vmHardDiskCollection = @() # VM Disks worksheet collection
    $ESXiCollection = @() # Host worksheet collection
    $ESXiNICCollection = @() # Host NICs worksheet collection
    $datastoreCollection = @() # Datastores worksheet collection
    $snapshotCollection = @() # Snapshots worksheet collection
    $rdmCollection = @() # RDMs worksheet collection
    $vmkCollection = @() # ESXi vmks
    $vssportgroupCollection = @() # Standard vSwitch portgroup collection
    $vdsportgroupCollection = @() # vDS portgroup collection
    $vcCollection = @() # vCenter collection
    $clusterCollection = @()
    $CDPCollection = @()


    # Now for the work - work against VC 
    foreach ($vc in $vCenter) {
        Write-Host "Connecting to vCenter $vCenter`n" -ForegroundColor Green
        Try {
            Connect-VIServer $vCenter -ErrorAction Stop
        }
        Catch {
            Write-Host "Unable to connect to the vCenter"
            Read-Host "Press ENTER to exit"
            Exit
        }
        Write-Host "Starting audit. This may take a few minutes to complete.`n" -ForegroundColor Green
    
        $VCName = $global:DefaultVIServer.name
    
        $vcInfo = [PSCustomObject]@{
            "vCenter" = $global:DefaultVIServer.Name
            "Version" = $global:DefaultVIServer.version
            "Build"   = $global:DefaultVIServer.build
        }
        $vcCollection += $vcInfo

        If ($Clusters -eq "Yes") {
        Write-Host "`nProcessing cluster infomation." -ForegroundColor Green
            forEach ($clus in Get-Cluster) {
                $clusterInfo = [PSCustomObject]@{
                    "Cluster"               = $clus.name
                    "Overall status"        = $clus.ExtensionData.Configstatus
                    "Configuration status"  = $clus.ExtensionData.Overallstatus
                    "HA Enabled"            = $clus.HAEnabled  
                    "HA Admission control"  = $clus.HAAdmissionControlEnabled
                    "HA restart priority "  = $clus.HARestartPriority
                    "HA isolation response" = $clus.HAIsolationResponse
                    "DRS"                   = $clus.DrsEnabled
                    "DRS mode"              = $clus.DrsMode
                    "DRS automation level"  = $clus.DrsAutomationLevel
                    "Number vMotions"       = $clus.ExtensionData.Summary.NumvMotions
                    "vSAN enabled"          = $clus.VsanEnabled
                    "Number hosts"          = $clus.ExtensionData.Summary.NumHosts
                    "CPU Cores"             = $clus.ExtensionData.Summary.Numcpucores
                    "CPU threads"           = $clus.ExtensionData.Summary.Numcputhreads
                    "VM Count"              = $clus.ExtensionData.Summary.usageSummary.Totalvmcount
                    "Powered off VM count"  = $clus.ExtensionData.Summary.usageSummary.Poweredoffvmcount
                }
                $clusterCollection += $clusterInfo
            }
        } # end If ($Clusters -eq "Yes") 

        $dcs = Get-Datacenter 
        foreach ($dc in $dcs) {
            # Gather the information for the Datastores worksheet              
            If ($Datastores -eq "Yes") {
                Write-Host "`nProcessing Datastores information in datacenter : $dc." -ForegroundColor Green
                $allDatastores = Get-Datastore -Location $dc
                foreach ($ds in $allDatastores) {
                    $DSDetails = $ds | Select-Object name, 
                        @{n = "Capacity"; E = { [math]::round($_.CapacityGB) } }, 
                        @{n = "FreeSpace"; E = { [math]::round($_.FreeSpaceGB) } }, 
                        @{N = "PercentFree"; E = { [math]::round($_.FreeSpaceGB / $_.CapacityGB * 100) } },
                        @{n = "ProvisionedGB";E= { [Math]::Round(($_.ExtensionData.Summary.Capacity - $_.ExtensionData.Summary.FreeSpace + $_.ExtensionData.Summary.Uncommitted)/1GB,0) } }

            
                    If ($ReportType -eq "Summary") {
                        $DSinfo = [PSCustomObject]@{
                            vCenter        = $VCNAme
                            DatastoreName  = $DSDetails.name
                            CapacityGB     = $DSDetails.Capacity
                            FreeSpaceGB    = $DSDetails.FreeSpace
                            PercentageFree = $DSDetails.PercentFree
                        } # end $DSinfo = [PSCustomObject]@
                    } else {
                        $DSinfo = [PSCustomObject]@{
                            vCenter        = $VCNAme
                            DatastoreName  = $DSDetails.name
                            State          = $ds.State
                            "SIOC Enabled" = $ds.StorageIOControlEnabled
                            "VMFS Version" = $ds.FileSystemVersion
                            NAA            = $ds.ExtensionData.Info.Vmfs.Extent.DiskName
                            CapacityGB     = $DSDetails.Capacity
                            FreeSpaceGB    = $DSDetails.FreeSpace
                            PercentageFree = $DSDetails.PercentFree
                            ProvisionedSpaceGB  = $DSDetails.ProvisionedGB
                        } # end $DSinfo = [PSCustomObject]@
                    } # end If ($ReportType -eq "Summary")
                    $datastoreCollection += $DSinfo       
                } # end foreach ($ds in $allDatastores)   
            }   # end If ($Datastores -eq "Yes")

            # Gather the information for the Hosts worksheet              
            If ($Hosts -eq "Yes") {
                Write-Host "`nProcessing Hosts information in datacenter : $dc." -ForegroundColor Green
                $allESXiHosts = Get-VMHost -Location $dc

                foreach ($ESXiHost in $allESXiHosts) {        

                    $VMsPerHost = $ESXiHost | Get-VM
                    if ($ESXiHost.IsStandalone) { $clusterName = 'Standalone' } else { $clusterName = $ESXiHost.Parent.Name }				
                        
                    # Calculate the host maximum, minimum, and average CPU and memory usage over the last 30 days ($TP variable)
                    # Not really sure of the value of this, but it's been requested in the past ...
                    $hoststat = "" | Select-Object HostName, MemMax, MemAvg, MemMin, CPUMax, CPUAvg, CPUMin
                    $statcpu = Get-Stat -Entity ($ESXiHost)-start (get-date).AddDays($TP) -Finish (Get-Date)-MaxSamples 10000 -stat cpu.usage.average
                    $statmem = Get-Stat -Entity ($ESXiHost)-start (get-date).AddDays($TP) -Finish (Get-Date)-MaxSamples 10000 -stat mem.usage.average
        
                    $cpu = $statcpu | Measure-Object -Property value -Average -Maximum -Minimum
                    $mem = $statmem | Measure-Object -Property value -Average -Maximum -Minimum
                        
                    $hoststat.CPUMax = [math]::round($cpu.Maximum, 2)
                    $hoststat.CPUAvg = [math]::round($cpu.Average, 2)
                    $hoststat.CPUMin = [math]::round($cpu.Minimum, 2)
                    $hoststat.MemMax = [math]::round($mem.Maximum, 2)
                    $hoststat.MemAvg = [math]::round($mem.Average, 2)
                    $hoststat.MemMin = [math]::round($mem.Minimum, 2)
                        
                    # Calculate the host uptime
                    $Uptime = $Esxihost | Select-Object @{N = "Uptime"; E = { New-Timespan -Start $_.ExtensionData.Summary.Runtime.BootTime -End (Get-Date) | Select-Object -ExpandProperty Days } }
                    $hostUptime = $Uptime.uptime

                    If ($ReportType -eq "Summary") {
                        $ESXinfo = [PSCustomObject]@{
                            vCenter           = $vcName
                            DC                = $dc.name
                            Cluster           = $clusterName
                            Hypervisor        = $ESXiHost.Name
                            ConnectionState   = $ESXiHost.ConnectionState
                            "Boot Time"       = $ESXiHost.ExtensionData.Summary.Runtime.BootTime
                            "Uptime (days)"   = $hostUptime
                            Version           = $ESXiHost.Version
                            CpuSockets        = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuPkgs
                            CpuCores          = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuCores
                            CpuThreads        = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuThreads
                            MemoryTotalGB     = $ESXiHost.MemoryTotalGB
                            NumVMs            = $allVMs.Count
                            "30 days Max CPU" = $hoststat.CPUMax
                            "30 days Min CPU" = $hoststat.CPUMin
                            "30 days Avg CPU" = $hoststat.CPUAvg
                            "30 days Max Mem" = $hoststat.MemMax
                            "30 days Min Mem" = $hoststat.MemMin
                            "30 days Avg Mem" = $hoststat.MemAvg
                        }
                    } else {
                        $ESXinfo = [PSCustomObject]@{
                            vCenter           = $vcName
                            DC                = $dc.name
                            Cluster           = $clusterName
                            Host              = $ESXiHost.Name
                            ConnectionState   = $ESXiHost.ConnectionState
                            "Boot Time"       = $ESXiHost.ExtensionData.Summary.Runtime.BootTime
                            "Uptime (days)"   = $hostUptime
                            Vendor            = $ESXiHost.ExtensionData.Summary.Hardware.Vendor
                            Model             = $ESXiHost.ExtensionData.Summary.Hardware.Model
                            Version           = $ESXiHost.Version
                            Build             = $ESXiHost.Build
                            CpuModel          = $ESXiHost.ExtensionData.Summary.Hardware.CpuModel
                            CpuSockets        = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuPkgs
                            CpuCores          = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuCores
                            CpuThreads        = $ESXiHost.ExtensionData.Summary.Hardware.NumCpuThreads
                            MemoryTotalGB     = $ESXiHost.MemoryTotalGB
                            NumVMs            = $VMsPerHost.Count
                            "30 days Max CPU" = $hoststat.CPUMax
                            "30 days Min CPU" = $hoststat.CPUMin
                            "30 days Avg CPU" = $hoststat.CPUAvg
                            "30 days Max Mem" = $hoststat.MemMax
                            "30 days Min Mem" = $hoststat.MemMin
                            "30 days Avg Mem" = $hoststat.MemAvg
                        } # end $ESXinfo = [PSCustomObject]@     
                        
                        # Gather host NIC details and we'll populate in a separate worksheet
                        $HostNICDetails = (Get-ESXcli -V2 -VMHost $ESXiHost).network.nic.list.Invoke()
                        foreach ($HostNic in $HostNICDetails) {
                            $ESXiNICInfo = [PSCustomObject]@{
                                Host          = $ESXiHost.Name
                                "NIC Name"    = $HostNIC.Name
                                "NIC MAC"     = $HostNIC.MACAddress
                                Description   = $HostNIC.Description
                                "Link status" = $HostNIC.Link
                                "Link Speed"  = $HostNIC.Speed
                                "Driver type" = $HostNIC.Driver
                                "MTU"         = $HostNIC.mtu 

                            }
                            $ESXINICCollection += $ESXiNICInfo                                 
                        } # End foreach ($HostNic in $HostNICDetails)

                        # Gather vmk info
                        $hostvmk = Get-VMHostNetworkAdapter -VMHost $ESXiHost | Where-Object {$_.name -like '*vmk*'} -ErrorAction SilentlyContinue
                        forEach ($vmk in $hostvmk) {
                            $vmkInfo = [PSCustomObject]@{
                                Host          = $vmk.VMHost
                                'vmk Name'    = $vmk.Name
                                IP            = $vmk.IP
                                'Subnet mask' = $vmk.subnetMask
                                MAC           = $vmk.Mac 
                                'Portgroup'   = $vmk.portgroupname
                                MTU           = $vmk.mtu 
                                "Management"  = $vmk.ManagementTrafficEnabled
                                "vMotion"     = $vmk.vMotionEnabled
                                "FT"          = $vmk.FaultToleranceLoggingEnabled
                                "VSAN"        = $vmk.VsanTrafficEnabled
                            }
                            $vmkCollection += $vmkInfo
                        } # end forEach ($vmk in $hostvmk) 

                        # Gather portgroup info for standard switches
                        foreach ($vsspg in $ESXiHost | get-virtualswitch -standard | get-virtualportgroup -Standard) {
                            $vsspgInfo = [PSCustomObject]@{
                                "Host"            = $ESXiHost.name
                                "Portgroup Name"  = $vsspg.name
                                "Virtual switch"  = $vsspg.virtualswitch
                                "VLAN ID"         = $vsspg.vlanid    
                            } 
                            $vssportgroupCollection += $vsspginfo
                        }

                        # Gather portgroup info for distributed switches
                        foreach ($vdspg in $ESXiHost | get-vdswitch | Get-VDPortgroup) {
                            If ($vdspg.name -like "*uplink*") {
                                $uplink = "Uplink"
                                $vlanID = ""
                            } else {
                                $vlanID = $vdspg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
                                $uplink = ""
                            }
                            
                            If ($vdspg.ExtensionData.Config.DefaultPortConfig.Vlan -is [VMware.Vim.VmwareDistributedVirtualSwitchPvlanSpec]){
                                $pvlanID = $vdspg.ExtensionData.Config.DefaultPortConfig.Vlan.PvlanId
                            } else {
                                $pvlanID = ""
                            }

                            $vdspgInfo = [PSCustomObject]@{
                                "Host"            = $ESXiHost.name
                                "Portgroup Name"  = $vdspg.name
                                "Virtual switch"  = $vdspg.vdSwitch.Name
                                "Uplink"          = $uplink
                                "VLAN ID"         = $vlanID
                                "PVLAN ID"        = $pvlanID
                            }
                            $vdsportgroupCollection += $vdspginfo
                        }
                        
                        # Gather CDP information if it's available - https://kb.vmware.com/s/article/1007069
                        $esxiHost | ForEach-Object { Get-View $_.ID } |
                        ForEach-Object { $esxname = $_.Name; Get-View $_.ConfigManager.NetworkSystem } |
                        ForEach-Object { foreach ($physnic in $_.NetworkInfo.Pnic) {
                                $CDPobject = "" | Select-Object Hostname, VMNIC, Speed, PCI, Port, NetworkSwitch, CDP_Address, VLAN, SwitchType

                                $pnicInfo = $_.QueryNetworkHint($physnic.Device)
                                foreach ($hint in $pnicInfo) {
                                    $CDPobject.Hostname = $esxname
                                    $CDPobject.VMNIC = $physnic.Device
                                    $CDPobject.Speed = $physnic.LinkSpeed.SpeedMb
                                    $CDPobject.PCI = $physnic.PCI
                                    if ( $hint.ConnectedSwitchPort ) {
                                        $CDPobject.Port = $hint.ConnectedSwitchPort.PortId
                                    }
                                    else {
                                        $CDPobject.Port = "No CDP information available."
                                    }
                                    $CDPobject.NetworkSwitch = $hint.ConnectedSwitchPort.DevId
                                    $CDPobject.CDP_Address = $hint.ConnectedSwitchPort.Address
                                    $CDPobject.VLAN = $hint.ConnectedSwitchPort.VLAN
                                    $CDPobject.SwitchType = $hint.ConnectedSwitchPort.HardwarePlatform
                                }
                                $CDPCollection += $CDPobject
                            }
                        } # end ForEach-Object { foreach ...

                    } # end If (ReportType -eq "Summary")
                    $ESXiCollection += $ESXinfo

                } # end foreach ($ESXiHost in $allESXiHosts)
            } # end If ($Hosts -eq "Yes")

            # Gather the information for the VMs worksheet              
            If ($VMs -eq "Yes") {
                Write-Host "`nProcessing VMs information in datacenter : $dc." -ForegroundColor Green

                $allVMs = get-vm -Location $dc
                foreach ($vm in $allVMs) {

                    $ESXiHost = $vm | Get-VMHost    
                    if ($ESXiHost.IsStandalone) { $clusterName = 'Standalone' } else { $clusterName = $ESXiHost.Parent.Name }				
                    
                    $ClusTemp = $vm | Get-VMHost
                    $clusterName = $ClusTemp.Parent.Name


                    # Collect the following only if $ReportType is detailed
                    If ($ReportType -eq "Detailed") {
                        # List the size of individual hard disks
                        $vmDisks = $vm | Get-HardDisk
                        $calcDisks = foreach ($disk in $vmDisks) {
                            "{0} = {1}" -f ($disk.Name), ($disk.CapacityGB)
                        }
                        $HardDiskSizes = $calcDisks -join "`n"

                        # Grab the IPs listed in the VM - requires VMwareTools be running.
                        # For those with multiple IPs, split them with `n
                        # Later, we will use ExportExcel to wrap the cells and top align the other cells in the worksheet
                        $ipTemp = $vm | Select-Object @{N = "IP Address"; E = { @($_.guest.IPAddress -join "`n") } } # pull all the IPs that VMwareTools will tell us about.
                        $ipList = $ipTemp | Select-Object -ExpandProperty "IP Address" # Drop the property name so we just have the IPs

                        # Get network adapter properties for MAC, connection state, type and the connected portgroup
                        # This will go into 4 cells per VM. Where there are multiple entries for a VM per cell
                        # they will be split into seperate lines within the cell
                        $vmNICS = $vm | Get-NetworkAdapter
                        $calcMACS = foreach ($nic in $vmNICS) {
                            "{0} : {1}" -f ($nic.Name), ($nic.MacAddress)
                        }
                        $Macs = $calcMACS -join "`n"

                        $calcNICState = foreach ($nic in $vmNICS) {
                            "{0} : {1}" -f ($nic.Name), ($nic.ConnectionState.Connected)
                        }
                        $NICState = $calcNICState -join "`n"

                        $calcNICType = foreach ($nic in $vmNICS) {
                            "{0} : {1}" -f ($nic.Name), ($nic.Type)
                        }
                        $NICType = $calcNICType -join "`n"

                        $calcPG = foreach ($nic in $vmNICS) {
                            "{0} : {1}" -f ($nic.Name), ($nic.NetworkName)
                        }
                        $PG = $calcPG -join "`n"

                        # List the size of the individual hard disks.
                        # If the VM has more than 1 disk, all disks are shown in a single cell, with each disk on a separate line within that cell.  
                        $vmDisks = $vm | Get-HardDisk
                        $calcDisks = foreach ($disk in $vmDisks) {
                            "{0} = {1}" -f ($disk.Name), ($disk.CapacityGB)
                        }
                        $HardDiskSizes = $calcDisks -Join ("`n")
                        
                        # Gather OS disk usage if VMwareTools are running
                        $OSUsage = $VM.Guest.Disks | Select-Object Path, 
                            @{n="Capacity"; E={[math]::round($_.CapacityGB)}}, 
                            @{n="FreeSpace"; E={[math]::round($_.FreeSpaceGB)}} 
                        $OSDisk = foreach ($partition in $OSUsage) {
                            "{0} Capacity = {1} Free = {2}" -f ($partition.Path), ($partition.Capacity),($partition.FreeSpace)
                        }
                        $Partitions = $OSDisk -join "`n"

                        # Gather RDM information    
                        $rdmList = $vm| Get-HardDisk | Where-Object {$_.DiskType -like "Raw*"} | Select-Object @{N="VM";E={$_.Parent}},
                            Name, DiskType,
                            @{N="NAA";E={$_.ScsiCanonicalName}},
                            @{N="VML";E={$_.DeviceName}},Filename,CapacityGB

                    } # end If $ReportType -eq "Detailed"

                    If ($ReportType -eq "Summary") {

                        $VMinfo = [PSCustomObject]@{
                            vCenter            = $VCName
                            DC                 = $dc.name
                            Cluster            = $clusterName
                            VM                 = $vm.Name
                            PowerState         = $vm.PowerState
                            NumCpu             = $vm.NumCpu
                            MemoryGB           = $vm.MemoryGB
                            ProvisionedSpaceGB = [Math]::Round($vm.ProvisionedSpaceGB, 2)
                            UsedSpaceGB        = [Math]::Round($vm.UsedSpaceGB, 2) # Used space on the datastore.
                            GuestOsFullName    = $vm.ExtensionData.Summary.Guest.GuestFullName
                            GuestName          = $vm.ExtensionData.Guest.Hostname
                        } # end $VMinfo = [PSCustomObject]@  
                    } else {
                        $VMinfo = [PSCustomObject]@{
                            vCenter                              = $VCName
                            DC                                   = $dc.name
                            Cluster                              = $clusterName
                            VM                                   = $vm.Name
                            PowerState                           = $vm.PowerState
                            NumCpu                               = $vm.NumCpu
                            MemoryGB                             = $vm.MemoryGB
                            ProvisionedSpaceGB                   = [Math]::Round($vm.ProvisionedSpaceGB, 2)
                            UsedSpaceGB                          = [Math]::Round($vm.UsedSpaceGB, 2) # Used space on the datastore.
                            "Individual Hard Disk sizes (GB)"    = $HardDiskSizes
                            "Volume/Partition : Capacity : Free" = $Partitions
                            GuestId                              = $vm.ExtensionData.Summary.Guest.GuestId
                            GuestOsFullName                      = $vm.ExtensionData.Summary.Guest.GuestFullName
                            GuestName                            = $vm.ExtensionData.Guest.Hostname
                            ToolsStatus                          = $vm.ExtensionData.Summary.Guest.ToolsStatus
                            "IP Address(es)"                     = $ipList
                            "MAC Addresses"                      = $Macs
                            "NIC Connection State"               = $NICState
                            "NIC Type"                           = $NICType
                            "Portgroup name"                     = $pg
                            ToolsVersion                         = $vm.ExtensionData.Config.Tools.ToolsVersion
                            MemoryHotAdd                         = $vm.ExtensionData.Config.MemoryHotAddEnabled
                            CPUHotAdd                            = $vm.ExtensionData.Config.CPUHotAddEnabled
                            "VM Hardware version"                = $vm.ExtensionData.Config.Version
                            Host                                 = $ESXiHost.name
                            Version                              = $ESXiHost.Version
                            Build                                = $ESXiHost.Build    
                        }   

                        If ($Performance -eq "Yes") {
                            # Calculate 30 day averages for memory, cpu, network and disk metrics, only for VMs powered on - otherwise set value to 0
                            # If we don't, we'll see lots of errors as it reports that it can't get the metric, and the respective cells in the worksheet
                            # would be left blank.
                            If ($vm.powerstate -eq "PoweredOn") {
                                $vmCPU = [Math]::Round(($vm | Get-Stat -Stat cpu.usagemhz.average -Start (Get-Date).AddDays($TP) -IntervalMins 5 | Measure-Object Value -Average).Average, 2)
                                $vmMem = [Math]::Round(($vm | Get-Stat -Stat mem.usage.average -Start (Get-Date).AddDays($TP) -IntervalMins 5 | Measure-Object Value -Average).Average, 2)
                                $vmNet = [Math]::Round(($vm | Get-Stat -Stat net.usage.average -Start (Get-Date).AddDays($TP) -IntervalMins 5 | Measure-Object Value -Average).Average, 2)
                                $vmDisk = [Math]::Round(($vm | Get-Stat -Stat disk.usage.average -Start (Get-Date).AddDays($TP) -IntervalMins 5 | Measure-Object Value -Average).Average, 2)
                            } else {
                                $vmCPU = 0
                                $vmMem = 0
                                $vmNet = 0
                                $vmDisk = 0
                            }

                            # Put VM performance related metrics into a separate custom object, so that we can populate a separate worksheet
                            $VMPerfInfo = [PSCustomObject]@{
                                VM                         = $vm.Name
                                PowerState                 = $vm.PowerState
                                MemoryReservation          = $vm.ExtensionData.ResourceConfig.MemoryAllocation.Reservation
                                MemoryLimit                = $vm.ExtensionData.ResourceConfig.MemoryAllocation.Limit
                                CPUReservation             = $vm.ExtensionData.ResourceConfig.CPUAllocation.Reservation
                                CPULimit                   = $vm.ExtensionData.ResourceConfig.CPUAllocation.Limit
                                Ballooning                 = $vm.ExtensionData.Summary.QuickStats.BalloonedMemory
                                "Avg CPU Usage (Mhz)"      = $vmCPU
                                "Avg Memory Usage (%)"     = $vmMem
                                "Avg Network Usage (KBps)" = $vmNet
                                "Avg Disk Usage (KBps)"    = $vmDisk
                            }  
                        } # end If ($Performance -eq "Yes")

                        # Create customm object for each harddisk on the VM giving its size and the datastore in which it resides.
                        # This will all go in the VM disks worksheet
                        foreach ($hd in $vmdisks) {
                            $vmHardDiskInfo = [PSCustomObject]@{
                                VM              = $vm.Name
                                HardDisk        = $hd.Name
                                Datastore       = $hd.FileName.Split(']')[0].TrimStart('[')
                                "Size (GB)"     = $hd.capacityGB 
                            } # end $vmHardDiskInfo = [PSCustomObject]@
                            $vmHardDiskCollection += $vmHardDiskInfo
                        } #end foreach ($hd in $vmdisks)
                    } # end If ($ReportType -eq "Summary")

                    # RDM worksheet
                    ForEach ($rdm in $rdmList) {
                        $rdmInfo = [PSCustomObject]@{
                            "VM Name"   = $rdm.VM
                            "Disk Name" = $rdm.Name
                            "Disk Type" = $rdm.DiskType
                            NAA         = $rdm.Naa
                            VML         = $rdm.VML
                            Filename    = $rdm.Filename
                            Capacity    = $rdm.CapacityGB
                        }
                        $rdmCollection += $rdmInfo
                    }

                    # Based on code from somewhere - not sure where now to attribute to.
                    # 2 sets = VM affinitiy/anti-affinity rules and VM to host affinity
                    # First, the VM affinity / anti-affinity
                    $DRSRules1 = Get-Cluster | Get-DrsRule
                    $VMAffinityDRSRules = ForEach ($DRSRule in $DRSRules1) {
                        "" | Select-Object -Property @{N = "Cluster"; E = { (Get-View -Id $DRSRule.Cluster.Id).Name } },
                            @{N = "Rule Name"; E = { $DRSRule.Name } },
                            @{N = "Enabled"; E = { $DRSRule.Enabled } },
                            @{N = "Rule Type"; E = { $DRSRule.Type } }, 
                            @{N = "VMs"; E = { $VMIds = $DRSRule.VMIds -split "," 
                                $TVMs = ForEach ($VMId in $VMIds) { 
                                    (Get-View -Id $VMId).Name
                                } 
                                $TVMs -join "," }
                            }
    
                    }
                    # Now the VM to host affinity
                    $DRSRules2 = Get-Cluster | Get-DrsRule -Type VMHostAffinity
                    $VMHostAffinityDRSRules = ForEach ($DRSRule in $DRSRules2) {
                        "" | Select-Object -Property @{N = "Cluster"; E = { (Get-View -Id $DRSRule.Cluster.Id).Name } },
                            @{N = "Rule Name"; E = { $DRSRule.Name } },
                            @{N = "Enabled"; E = { $DRSRule.Enabled } },
                            @{N = "Rule Type"; E = { $DRSRule.Type } }, 
                            @{N = "VMs"; E = { $VMIds = $DRSRule.VMIds -split "," 
                                $VMs = ForEach ($VMId in $VMIds) { 
                                    (Get-View -Id $VMId).Name
                                } 
                                $VMs -join "," }
                            }
                    } 
                    # ugh - combine the 2 sets of results 
                    $AllDRSRules = $VMAffinityDRSRules + $VMHostAffinityDRSRules

                    $VMCollection += $VMinfo
                    $VMPerfCollection += $VMPerfInfo

                } # end foreach ($vm in $allVMs)


            } # end If ($VMs -eq "Yes")

            # Gather the information for the Snapshots worksheet              
            If ($Snapshots -eq "Yes") {
                Write-Host "`nProcessing Snapshots information in datacenter : $dc." -ForegroundColor Green
                foreach ($snap in Get-VM -Location $dc | Get-Snapshot) {
                    $ds = Get-Datastore -VM $snap.vm
                    $SnapshotAge = ((Get-Date) - $snap.Created).Days

                        # who's the culprit?
                        $snapshotEvent = Get-VIEvent -Entity $snap.VM -Types Info -Finish $snap.Created.AddMinutes(10) | Where-Object {$_.FullFormattedMessage -imatch 'Task: Create virtual machine snapshot'}
                        if ($snapshotEvent -ne $null) {
                            If ($snapshotEvent.count -gt 1) {
                                # Just in case there are multiple events, get the latest one.
                                $SnapshotEvent = $SnapshotEvent | Sort-Object -Property CreatedTime | Select-Object -Last 1 
                            }
                            $CreatedBy = $SnapshotEvent.Username
                        } else {
                            $CreatedBy = "Unknown" # ok, can't figure an entry - maybe the snap is so old it's rotated out, or maybe the script logic has failed.
                        }
                    
                    $snapinfo = [PSCustomObject]@{
                        "vCenter"                   = $vcName
                        "VM"                        = $snap.vm
                        "Snapshot Name"             = $snap.name
                        "Description"               = $snap.description
                        "Created"                   = $snap.created
                        "Created by"                = $CreatedBy
                        "Snapshot age (days)"       = $SnapshotAge
                        "Snapshot size (GB)"        = [math]::round($snap.sizeGB)
                        "Datastore"                 = $ds[0].name
                        "Datastore free space (GB)" = [math]::round($ds[0].FreeSpaceGB)
                    } # end $snapinfo = [PSCustomObject]@
                    $snapshotCollection += $snapinfo
                } # end foreach ($snap in Get-VM | Get-Snapshot)
            } # end If ($Snapshots -eq "Yes")

        } # end foreach ($dc in $dcs)

        Disconnect-VIServer -Server * -Force -Confirm:$false
    } # end foreach ($vc in $vcs)
        
    # Now start preparing to export the results to file(s)
    $date = Get-Date -Format "yyyy-MMM-dd-HHmmss"
            
    If ($OutputFormat -eq "Excel") {
        # Generate the worksheets in the .xlsx - relies on the ImportExcel module from the start of the script.
        $xlsx_output_file = "$script_dir\$VCName-Audit-$date.xlsx"
        $vcCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "VC" -Autosize

        If ($Clusters -eq "Yes") {
            $clusterCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "Clusters" -Autosize
        }

        If ($vms -eq "Yes") { 
            If ($ReportType -eq "Detailed") {
                # For the VMs worksheet, we need to do some extra work with the IPAddress(es) column to
                # wrap the text in the cell - earlier we split the IPs onto newlines.
                # Also then doing a vertical alignment of the other columns to top - this is my preference.
                $a = $VMCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname VMs -AutoSize -PassThru
                $a.workbook.Worksheets["VMs"].Column(10).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(11).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(16).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(17).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(18).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(19).Style.Wraptext = $true
                $a.workbook.Worksheets["VMs"].Column(20).Style.Wraptext = $true
                foreach ($c in 1..27) {
                    # Set vertical alignment for each column to top. For ease, doing this for columns 1 -> 26. There will be a more "correct" way to do this.
                    $a.workbook.Worksheets["VMs"].Column($c).Style.VerticalAlignment = "Top"
                    $a.workbook.Worksheets["VMs"].Column($c).Style.HorizontalAlignment = "Left"
                }
                $a.save()
                $a.dispose()
                $vmHardDiskCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "VM Disks" -Autosize
                $rdmCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "RDMs" -Autosize
                $AllDRSRules | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "DRS Rules" -Autosize
                If ($Performance -eq "Yes") {
                    $VMPerfCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "VM Performance" -AutoSize
                }

            } else {
                $VMCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname VMs -AutoSize
            }
        }

        If ($Hosts -eq "Yes") { 
            If ($ReportType -eq "Detailed") {
                $ESXiCollection | Sort-Object -Property Hypervizor | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "ESXi hosts" -AutoSize             
                $ESXiNICCollection | Sort-Object -Property Host | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "ESXi NICs" -AutoSize     
                $vmkCollection | Sort-Object -Property Host | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "ESXi vmks" -AutoSize     
                $vssportgroupCollection | Sort-Object -Property Host | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "vss Portgroups" -AutoSize
                $vdsportgroupCollection | Sort-Object -Property Host | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "vDS Portgroups" -AutoSize
                $CDPCollection | Sort-Object -Property Host | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "CDP" -AutoSize     
            } else {
                $ESXiCollection | Sort-Object -Property Hypervizor | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "ESXi hosts" -AutoSize             
            }
        }

        If ($Datastores -eq "Yes") { 
            $datastoreCollection | Sort-Object -Property DatastoreName | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname Datastores -AutoSize 
        }

        If ($Snapshots -eq "Yes") { 
            $snapshotCollection | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname Snapshots -AutoSize 
        }

        Write-Host "`nAudit generated in $xlsx_output_file" -ForegroundColor Green
    } else {
        # Output form must be .csv so generate files for each that were selected.
        If ($Clusters -eq "Yes") {
            $clusters_csv = "$script_dir\$VCName-Clusters-Audit-$date.csv" 
            $clusterCollection | Export-CSV -NoTypeInformation -Path $clusters_csv
            Write-Host "Clusters audit : $clusters_csv" -ForegroundColor Green    

        }

        If ($vms -eq "Yes") { 
            $vm_csv = "$script_dir\$VCName-VM-Audit-$date.csv" 
            $VMCollection | Export-CSV -NoTypeInformation -Path $VM_csv
            Write-Host "VM audit : $vm_csv" -ForegroundColor Green    
            If ($ReportType -eq "Detailed") {
                $vmDisks_csv = "$script_dir\$VCName-VMDisks-Audit-$date.csv" 
                $vmHardDiskCollection | Export-CSV -NoTypeInformation -Path $VMDisks_csv
                Write-Host "VM disks audit : $vmDisks_csv" -ForegroundColor Green  
                $rdm_csv = "$script_dir\$VCName-RDMs-Audit-$date.csv" 
                $rdmCollection | Export-CSV -NoTypeInformation -Path $rdm_csv
                Write-Host "RDMs audit : $rdm_csv" -ForegroundColor Green  
                If ($Performance -eq "Yes") {
                    $vmperf_csv = "$script_dir\$Customer-$VCName-$Ticket-VMPerf-Audit-$date.csv" 
                    $VMPerfCollection | Export-CSV -NoTypeInformation -Path $VMperf_csv
                    Write-Host "VM Performance audit : $vm_csv" -ForegroundColor Green      
                }
            }
        }
        If ($Hosts -eq "Yes") { 
            $ESXiHosts_csv = "$script_dir\$VCName-ESXi-Hosts-Audit-$date.csv" 
            $ESXiCollection | Export-CSV -NoTypeInformation -Path $ESXiHosts_csv   
            Write-Host "ESXi Hosts audit : $ESXiHosts_csv" -ForegroundColor Green    

            $ESXiNICS_csv = "$script_dir\$VCName-ESXi-Hosts-Audit-$date.csv" 
            $ESXiNICCollection | Export-CSV -NoTypeInformation -Path $ESXiNICS_csv   
            Write-Host "ESXi NICs audit : $ESXiNICS_csv" -ForegroundColor Green    

            $ESXivmk_csv = "$script_dir\$VCName-ESXi-vmks-Audit-$date.csv" 
            $vmkCollection | Export-CSV -NoTypeInformation -Path $ESXivmk_csv   
            Write-Host "ESXi vmks audit : $ESXivmk_csv" -ForegroundColor Green    

            $vssportgroups_csv = "$script_dir\$VCName-ESXi-vss-portgroups-Audit-$date.csv" 
            $vssportgroupCollection | Export-CSV -NoTypeInformation -Path $vssportgroups_csv   
            Write-Host "ESXi Hosts audit : $vssportgroups_csv" -ForegroundColor Green    

            $vdsportgroups_csv = "$script_dir\$VCName-ESXi-vds-portgroups-Audit-$date.csv" 
            $vdsportgroupCollection | Export-CSV -NoTypeInformation -Path $vdsportgroups_csv   
            Write-Host "ESXi Hosts audit : $vdsportgroups_csv" -ForegroundColor Green   
            
            $CDP_csv = "$script_dir\$Customer-$VCName-$Ticket-ESXi-CDP-Audit-$date.csv" 
            $CDPCollection | Export-CSV -NoTypeInformation -Path $CDP_csv   
            Write-Host "ESXi CDP audit : $CDP_csv" -ForegroundColor Green    


        }
        If ($Datastores -eq "Yes") { 
            $Datastore_csv = "$script_dir\$VCName-Datastore-Audit-$date.csv" 
            $datastoreCollection | Export-CSV -NoTypeInformation -Path $Datastore_csv 
            Write-Host "Datastore audit : $Datastore_csv" -ForegroundColor Green    
        }
        If ($Snapshots -eq "Yes") { 
            $snapshot_csv = "$script_dir\$VCName-Snapshot-Audit-$date.csv"
            $snapshotCollection | Export-CSV -NoTypeInformation -Path $snapshot_csv 
            Write-Host "Snapshot audit : $snapshot_csv" -ForegroundColor Green        
        }    
    } # end If ($OutputFormat -eq "Excel")   

} # end Get-EnvironmentReport
    