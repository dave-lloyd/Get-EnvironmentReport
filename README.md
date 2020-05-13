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

### VMs - Information included :
    VM name
    Powerstate
    CPUs
    Memory
    Provisioned space
    Guest OS full name
    Guest Name - guestname at the OS level - requires that VMwareTools be running.
    IP Address(es) - requires VMwareTools be running. Devices with multiple IPs will report all IPs in a cell, 1 IP per line. If you have VMs with multiple VMs, this can look inelegant.
    VMware Tools status
    VMware Tools version
    CPU hot add - True or False
    Memory hot add - True or False
    VM hardware version
    Host VM is running on
    Host version number
    Host build number

### ESXi hosts - Information included :
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

### Datastores - Information included :
    Datastore name
    NAA
    Capacity
    Free space
    Percentage free - values are rounded up.

### Snapshots - Information included :
    VM
    Snapshot name
    Description
    Date created
    Snapshot age
    Snapshot size
    Datastore
    Datastore free space

## Credentials required
An account with the Read-only role is sufficient for this script.

## Output formats
You will have the option to generate a single .xlsx file, which will have a worksheet per category, or individual .csv files
for each of the selected categories. The generation of the .xlsx file relies on the use of the ImportExcel Powershell module.
Select Excel or CSV for the OutputFormat parameter. IF the check for the Excel module fails, the script will fall back to
creating .csv files instead.

## Location of generated reports.
Reports will be generated in a folder in your Documents folder called "EnvironmentReports\Reports".

