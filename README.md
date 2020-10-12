# Get-EnvironmentReport
This script is a collection of functions designed to perform a basic audit of an environment. It is based on requests received in my work environment, where it's not permitted to use RVTools, so this was created in order to get a quick general audit for an environment. So some of the (derived) properties may not make sense generally, but are a reflection of information that have been requested for on occassion, and it's easier to have the information, and then remove it, than trying to "bolt on" more data for a report. 

Some code and functions has also been removed, mainly around authentication and parameters for the main function removed (as they tied more in with using a ticketing system.) 

Hopefully the structure allows for easily removing properties that aren't needed/useful, and also to add further properties which may be applicable.

When the script is run, it is necessary to specify whether or not to include information on VMs, Hosts, Datastores and Snapshot, and also whether to generate the audit as a single .xlsx file or multiple .csv files. To generate a single .xlsx file, the ImportExcel module is required - https://www.powershellgallery.com/packages/ImportExcel/7.1.0 

## Running the script
Download it, and then dot source it, so, in the folder you have the downloaded version :
- . .\Get-EnvironmentReport.ps1

Once that's done, you can run help in the normal way to get the parameters and some examples
Help Get-EnvironmentReport -full

## Options :
The core options are as follows. If you export as Excel, each of these will be a worksheet in the final spreadsheet. If you export as CSV, then these would be the individual CSV files :
- VMs
- ESXi Hosts
- Datastores
- Snapshots


### VMs - Information included for "Summary" report type :
    VM name
    Powerstate
    CPUs
    Memory 
    Provisioned space
    Guest OS full name 
    Guest Name - guestname at the OS level - requires that VMwareTools be running.

### VMs - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
    Individual Hard Disk sizes (GB) - the size of the individual harddisks as presented to the VM.
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

### VM Performance - only produced if "Detailed" report type is selected :
    VM Name
    Powerstate
    MemoryReservation - whether VM has a memory reservation and if so, what size. No reservation = 0
    MemoryLimit - whether VM has a memory limit and if so, what size. -1 = no limit.
    CPUReservation - whether VM has a CPU reservation and if so, what size. No reservation = 0
    CPULimit - whether VM has a CPU limit and if so, what size. -1 = no limit.
    Ballooning - value for the Summary.QuickStats.BalloonedMemory property
    Avg CPU Usage (Mhz) - derived from cpu.usagemhz.average metric over 30 days with 5 minute interval
    Avg Memory Usage (%) - derived from memory.usage.average metric over 30 days with 5 minute interval
    Avg Network Usage (KBps) - derived from network.usage.average metric over 30 days with 5 minute interval
    Avg Disk Usage (KBps) - derived from disk.usage.average metric over 30 days with 5 minute interval

You should probably treat this information with caution, given the sample rate and that it presents a single value as a 30 day representation of "performance."

### VM Disks - only produced if "Detailed" report type is selected :
    VM 
    HardDisk 
    Datastore 
    Size in GB

### RDMs - only produced if "Detailed" report type is selected :
    VM Name
    Disk Name
    Disk Type
    NAA
    VML         
    Filename   
    Capacity    

### ESXi hosts - Information included for "Summary" report type :
    Host name 
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

### ESXi hosts - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
    Vendor 
    Model  
    ESXi build number 
    CPU Model 

### ESXi NICs - Information included only if hosts AND "detailed" report type are selected :
    Host
    NIC Name
    MAC Address
    Description
    Link status
    Link Speed
    Driver Type
    MTU

### ESXi vmks - only produced if hosts AND "detailed" report type are selected.
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

### Datastores - Information included for "Summary" report type :
    Datastore name 
    Capacity 
    Free space 
    Percentage free - values are rounded up. 

### Datastores - Information included for "Detailed" report type - all the info in the Summary option, plus the following :
    Datastore name 
    State
    SIOC - whether enabled or not
    VMFS version
    NAA 
    ProvisionedGB - view to see if overprovisioned due to thin provisioning.

## Credentials required
An account with the Read-only role is sufficient for this script.

## Output formats
You will have the option to generate a single .xlsx file, which will have a worksheet per category, or individual .csv files
for each of the selected categories. The generation of the .xlsx file relies on the use of the ImportExcel Powershell module.
Select Excel or CSV for the OutputFormat parameter. IF the check for the Excel module fails, the script will fall back to
creating .csv files instead.

## Location of generated reports.
Reports will be generated in a folder in your Documents folder called "EnvironmentReports\Reports".

