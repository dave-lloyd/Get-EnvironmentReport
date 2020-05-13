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

         VMs - Information included :
            VM name
            Powerstate
            CPUs
            Memory
            Provisioned space
            Guest OS full name
            Guest Name - guestname at the OS level - requires that VMwareTools be running.
            IP Address(es) - requires VMwareTools be running. All IPs for a VM will be presented in a single cell - one per line.
            VMware Tools status
            VMware Tools version
            CPU hot add - True or False
            Memory hot add - True or False
            VM hardware version
            Host VM is running on
            Host version number
            Host build number

         ESXi hosts - Information included :
            Host name
            Connection state
            Host uptime in days
            Vendor
            Model
            ESXi version number
            ESXi build number
            CPU Model
            Number sockets
            Number of cores
            Total CPU threads
            Memory 
            Number of VMs running on the host
            30 day maximum | minimum | average for CPU and Memory

         Datastores - Information included :
            Datastore name
            NAA
            Capacity
            Free space
            Percentage free - values are rounded up.

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
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate an .xlsx file Named 10.10.10.10-Audit-<date>.xlsx which contains worksheets for VMs, ESXi hosts and Datastores
         Get-EnvironmentReport -vcs 10.10.10.10 -VMs Yes -Hosts Yes -Datastores Yes -Snapshots No -OutputFormat Excel 
    
        .EXAMPLE
         The following example will connect to VC 10.10.10.10 and generate a series of .csv files, one each containing an audit of VMs, ESXi hosts, Datastores and Snapshots in the environment. 
         Get-EnvironmentReport -vcs 10.10.10.10 -VMs Yes -Hosts Yes -Datastores Yes -Snapshots Yes -OutputFormat CSV 
              
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
        [string]$OutputFormat  
    )
    
    Set-InitializeAudit

    # Check to see if ImportExcel is available. If it's not, we'll have to default to .csv
    # We probably could/should offer to install the module instead - at least scoped for the current user.
    If ($OutputFormat -eq "Excel") {
        $ExcelAvailable = Test-ForImportExcel
        If (-not $ExcelAvailable) {
            Write-Host "ImportExcel appears to note be available. Switching output format to .csv" -ForegroundColor Green
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
    $ESXiCollection = @()
    $datastoreCollection = @()
    $snapshotCollection = @()
                
    # Now for the work - work against VC 
    foreach ($vc in $vCenter) {
        Write-Host "Connecting to vCenter $vCenter`n" -ForegroundColor Green
        Try {
            Connect-VIServer $vCenter -ErrorAction Stop
        } Catch {
            Write-Host "Unable to connect to the vCenter"
            Read-Host "Press ENTER to exit"
            Exit
        }
        Write-Host "Starting audit. This may take a few minutes to complete.`n" -ForegroundColor Green
    
        $VCName = $global:DefaultVIServer.name
    
        $dcs = Get-Datacenter 
        foreach ($dc in $dcs) {
            # Get the information for the Datastores worksheet    
            $allDatastores = Get-Datastore -Location $dc
            foreach ($ds in $allDatastores) {
                $DSDetails = $ds | Select-Object name, @{n = "Capacity"; E = { [math]::round($_.CapacityGB) } }, @{n = "FreeSpace"; E = { [math]::round($_.FreeSpaceGB) } }, @{N = "PercentFree"; E = { [math]::round($_.FreeSpaceGB / $_.CapacityGB * 100) } }
            
                $info = [PSCustomObject]@{
                    vCenter        = $VCNAme
                    DatastoreName  = $DSDetails.name
                    NAA            = $ds.ExtensionData.Info.Vmfs.Extent.DiskName
                    CapacityGB     = $DSDetails.Capacity
                    FreeSpaceGB    = $DSDetails.FreeSpace
                    PercentageFree = $DSDetails.PercentFree
                }
                $datastoreCollection += $info       
            } # end foreach ($ds in $allDatastores)   
    
            # Get the information for the ESXi hosts worksheet    
            $allESXiHosts = Get-VMHost -Location $dc
            foreach ($ESXiHost in $allESXiHosts) {                
                $allVMs = get-vm -Location $ESXiHost
                    
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

                $info = [PSCustomObject]@{
                    vCenter           = $vcName
                    DC                = $dc.name
                    Cluster           = $clusterName
                    "ESXi host"       = $ESXiHost.Name
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
                }
                $ESXiCollection += $info
    
                # Get the information for the VMs worksheet
                foreach ($vm in $allVMs) {
                    # Grab the IPs listed in the VM - requires VMwareTools be running.
                    # For those with multiple IPs, split them with `n
                    # Later, we will use ExportExcel to wrap the cells and top align the other cells in the worksheet
                    $ipTemp = $vm | Select-Object @{N = "IP Address"; E = { @($_.guest.IPAddress -join "`n") } } # pull all the IPs that VMwareTools will tell us about.
                    $ipList = $ipTemp | Select-Object -ExpandProperty "IP Address" # Drop the property name so we just have the IPs

                    $info = [PSCustomObject]@{
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
                    $VMCollection += $info
                } # end foreach ($vm in $allVMs)
            } # end foreach ($ESXiHost in $allESXiHosts)
    
            # Get the information for the Snapshots worksheet
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
                }
                $snapshotCollection += $snapinfo
            } # end foreach ($snap in Get-VM | Get-Snapshot)
        } # end foreach ($dc in $dcs)
    
        Disconnect-VIServer -Server * -Force -Confirm:$false
    } # end foreach ($vc in $vcs)
        
    # Now start preparing to export the results to file(s)
    $date = Get-Date -Format "yyyy-MMM-dd-HHmmss"
            
    If ($OutputFormat -eq "Excel") {
        # Generate the worksheets in the .xlsx - relies on the ImportExcel module from the start of the script.
        $xlsx_output_file = "$script_dir\$VCName-Audit-$date.xlsx"

        If ($vms -eq "Yes") { 
            # For the VMs worksheet, we need to do some extra work with the IPAddress(es) column to
            # wrap the text in the cell - earlier we split the IPs onto newlines.
            # Also then doing a vertical alignment of the other columns to top - this is my preference.
            $a = $VMCollection | Sort-Object -Property Cluster, VM | Export-Excel $xlsx_output_file -BoldTopRow -AutoFilter -FreezeTopRow -WorkSheetname VMs -AutoSize -PassThru
            $a.workbook.Worksheets["VMs"].Column(14).Style.Wraptext = $true
            foreach ($c in 1..26) {
                # Set vertical alignment for each column to top. For ease, doing this for columns 1 -> 26. 
                # There will be a more "correct" way to determine how many columns ...
                $a.workbook.Worksheets["VMs"].Column($c).Style.VerticalAlignment = "Top"
            }
            $a.save()
            $a.dispose()
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
    }
    else {
        # Output form must be .csv so generate files for each that were selected.
        If ($vms -eq "Yes") { 
            $vm_csv = "$script_dir\$VCName-VM-Audit-$date.csv" 
            $VMCollection | Export-CSV -NoTypeInformation -Path $VM_csv
            Write-Host "VM audit : $vm_csv" -ForegroundColor Green    
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
    }    
} # end Get-EnvironmentReport
    