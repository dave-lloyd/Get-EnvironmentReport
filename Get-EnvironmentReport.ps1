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
    }
    else {
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
    }
    else {
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

         VMs - Information included for "Summary" report type :
            VM name
            Powerstate
            CPUs
            Memory 
            Provisioned space
            Guest OS full name 
            Guest Name - guestname at the OS level - requires that VMwareTools be running.

         VMs - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
            IP Address(es) - requires VMwareTools be running. 
            Guest OS ID - VMware ID for the OS type   
            VMware Tools status
            VMware Tools version
            CPU hot add - True or False
            Memory hot add - True or False
            VM hardware version
            Host VM is running on
            Host version number
            Host build number

        VM Performance - only produced if "Detailed" report type is selected.
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

         ESXi hosts - Information included for "Summary" report type :
            Host name 
            Connection state 
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

         Snapshots - Information included :
            VM 
            Snapshot name 
            Description 
            Date created 
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

        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate an .xlsx file Named 10.10.10.10-Audit-<date>.xlsx which contains worksheets for VMs, ESXi hosts and Datastores
         Get-EnvironmentReport -vCenter 10.10.10.10 -VMs Yes -Hosts Yes -Datastores Yes -Snapshots No -OutputFormat Excel -ReportType Details
    
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate a series of .csv files, one each containing an audit of VMs, ESXi hosts, Datastores and Snapshots in the environment. 
         Get-EnvironmentReport -vCenter 10.10.10.10 -VMs Yes -Hosts Yes -Datastores Yes -Snapshots Yes -OutputFormat CSV -ReportType Summary
              
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
        [string]$ReportType  
       
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
    $VMCollection = @()
    $VMPerfCollection = @()
    $ESXiCollection = @()
    $datastoreCollection = @()
    $snapshotCollection = @()
                
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
    
        $dcs = Get-Datacenter 
        foreach ($dc in $dcs) {
            # Gather the information for the Datastores worksheet              
            If ($Datastores -eq "Yes") {
                Write-Host "`nProcessing Datastores information in datacenter : $dc." -ForegroundColor Green
                $allDatastores = Get-Datastore -Location $dc
                foreach ($ds in $allDatastores) {
                    $DSDetails = $ds | Select-Object name, @{n = "Capacity"; E = { [math]::round($_.CapacityGB) } }, @{n = "FreeSpace"; E = { [math]::round($_.FreeSpaceGB) } }, @{N = "PercentFree"; E = { [math]::round($_.FreeSpaceGB / $_.CapacityGB * 100) } }
            
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
                    #$allVMs = get-vm -Location $ESXiHost
                        
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
                            Hypervisor        = $ESXiHost.Name
                            ConnectionState   = $ESXiHost.ConnectionState
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
                            NumVMs            = $allVMs.Count
                            "30 days Max CPU" = $hoststat.CPUMax
                            "30 days Min CPU" = $hoststat.CPUMin
                            "30 days Avg CPU" = $hoststat.CPUAvg
                            "30 days Max Mem" = $hoststat.MemMax
                            "30 days Min Mem" = $hoststat.MemMin
                            "30 days Avg Mem" = $hoststat.MemAvg
                        } # end $ESXinfo = [PSCustomObject]@       
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
                    # Grab the IPs listed in the VM - requires VMwareTools be running.
                    # For those with multiple IPs, split them with `n
                    # Later, we will use ExportExcel to wrap the cells and top align the other cells in the worksheet
                    $ipTemp = $vm | Select-Object @{N = "IP Address"; E = { @($_.guest.IPAddress -join "`n") } } # pull all the IPs that VMwareTools will tell us about.
                    $ipList = $ipTemp | Select-Object -ExpandProperty "IP Address" # Drop the property name so we just have the IPs

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
                    } # end If ($vm.powerstate -eq "PoweredOn")

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
                            vCenter               = $VCName
                            DC                    = $dc.name
                            Cluster               = $clusterName
                            VM                    = $vm.Name
                            PowerState            = $vm.PowerState
                            NumCpu                = $vm.NumCpu
                            MemoryGB              = $vm.MemoryGB
                            ProvisionedSpaceGB    = [Math]::Round($vm.ProvisionedSpaceGB, 2)
                            UsedSpaceGB           = [Math]::Round($vm.UsedSpaceGB, 2) # Used space on the datastore.
                            GuestId               = $vm.ExtensionData.Summary.Guest.GuestId
                            GuestOsFullName       = $vm.ExtensionData.Summary.Guest.GuestFullName
                            GuestName             = $vm.ExtensionData.Guest.Hostname
                            ToolsStatus           = $vm.ExtensionData.Summary.Guest.ToolsStatus
                            "IP Address(es)"      = $ipList
                            ToolsVersion          = $vm.ExtensionData.Config.Tools.ToolsVersion
                            MemoryHotAdd          = $vm.ExtensionData.Config.MemoryHotAddEnabled
                            CPUHotAdd             = $vm.ExtensionData.Config.CPUHotAddEnabled
                            "VM Hardware version" = $vm.ExtensionData.Config.Version
                            Host                  = $ESXiHost.name
                            Version               = $ESXiHost.Version
                            Build                 = $ESXiHost.Build    
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
                        } # end $VMPerfInfo = [PSCustomObject]@
                    } # end If ($ReportType -eq "Summary")
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
                
                    $snapinfo = [PSCustomObject]@{
                        "vCenter"                   = $vcName
                        "VM"                        = $snap.vm
                        "Snapshot Name"             = $snap.name
                        "Description"               = $snap.description
                        "Created"                   = $snap.created
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

        If ($vms -eq "Yes") { 
            If ($ReportType -eq "Detailed") {
                # For the VMs worksheet, we need to do some extra work with the IPAddress(es) column to
                # wrap the text in the cell - earlier we split the IPs onto newlines.
                # Also then doing a vertical alignment of the other columns to top - this is my preference.
                $a = $VMCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname VMs -AutoSize -PassThru
                $a.workbook.Worksheets["VMs"].Column(14).Style.Wraptext = $true
                foreach ($c in 1..27) {
                    # Set vertical alignment for each column to top. For ease, doing this for columns 1 -> 26. There will be a more "correct" way to do this.
                    $a.workbook.Worksheets["VMs"].Column($c).Style.VerticalAlignment = "Top"
                }
                $a.save()
                $a.dispose()
                $VMPerfCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "VM Performance" -AutoSize

            } else {
                $VMCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname VMs -AutoSize
            }
        }

        If ($Hosts -eq "Yes") { 
            $ESXiCollection | Sort-Object -Property Hypervizor | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname "ESXi hosts" -AutoSize 
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
        If ($vms -eq "Yes") { 
            $vm_csv = "$script_dir\$VCName-VM-Audit-$date.csv" 
            $VMCollection | Export-CSV -NoTypeInformation -Path $VM_csv
            Write-Host "VM audit : $vm_csv" -ForegroundColor Green    
            If ($ReportType -eq "Detailed") {
                $vmperf_csv = "$script_dir\$VCName-VMPerf-Audit-$date.csv" 
                $VMPerfCollection | Export-CSV -NoTypeInformation -Path $VMPerf_csv
                Write-Host "VM audit : $vm_csv" -ForegroundColor Green                    
            }
        }
        If ($Hosts -eq "Yes") { 
            $ESXiHosts_csv = "$script_dir\$VCName-ESXi-Hosts-Audit-$date.csv" 
            $ESXiCollection | Export-CSV -NoTypeInformation -Path $ESXiHosts_csv   
            Write-Host "ESXi Hosts audit : $ESXiHosts_csv" -ForegroundColor Green    
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
    